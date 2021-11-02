using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Renci.SshNet.Common;
using Renci.SshNet.Sftp;

namespace Agent2._0
{
    class Program
    {
        static InfoServerDB dbSvInfo = new InfoServerDB();
        static InfoServerSftp sftpSvInfo = new InfoServerSftp();
        static ServerDatabase db;
        static ServerSftp svSftp;
        static DateTime now = DateTime.Now;
        static void Main(string[] args)
        {
            Console.Title = "Agent 2.0";
            // Init server
            db = new ServerDatabase(dbSvInfo);
            svSftp = new ServerSftp(sftpSvInfo);
            
            if (db.conn.State == System.Data.ConnectionState.Open)
            {
                Console.WriteLine("[Info] Connected to Database");
                LoadCampaigns();
            }
            else
            {
                Console.WriteLine("[Error] Unable to connect to database");
            }
            RunCalibStationAPI();
            Console.ReadLine();
        }

        static void LoadCampaigns()
        {
            List<tblCampaign> ActiveCampaignList = new List<tblCampaign>();
            MySqlDataReader rdr = db.Reader("SELECT id, created_date, last_updates, log_path FROM tbl_campaign");
            while (rdr.Read())
            {
                DateTime createTime = DateTime.ParseExact(rdr.GetString(1), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                DateTime lastUpdate = DateTime.ParseExact(rdr.GetString(2), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                if (DateTime.Compare(createTime, now) <= 0 && DateTime.Compare(now, lastUpdate) <= 0)
                {
                    Console.WriteLine("[Info] Found campaign: " + rdr.GetString(0) + "--" + rdr.GetString(1) + "--" + rdr.GetString(2) + "--" + rdr.GetString(3));
                    ActiveCampaignList.Add(new tblCampaign(rdr.GetInt32(0), createTime, lastUpdate, rdr.GetString(3)));
                }
            }
            rdr.Close();
            
            foreach(var campaign in ActiveCampaignList)
            {
                LoadCampaign(campaign);
            }
            
        }

        static void LoadCampaign(tblCampaign campaign)
        {
            //    "1.1 Do kiem TRX",
            //    "1.2 Do kiem PA",
            //    "1.3 Do kiem Filter",
            //    "1.4 Do kiem Antenna",
            //    "2. Lap rap RRU",
            //    "3.1 Luu Serial vao EEPROM RRU",
            //    "3.2 Do kiem Tx, Rx",
            //    "4. Test kin khi",
            //    "5. Burn in",
            //    "6. Do kiem TX RX sau burn in",
            //    "7. Test kin khi sau burn in",
            //    "8.1 OQC Chu trinh nhiet",
            //    "8.2 OQC Rung xoc",
            //    "9. Do kiem TX RX sau rung xoc, chu trinh nhiet",
            //    "10. Package"

            svSftp.Connect();
            if (svSftp.sftp.IsConnected)
            {
                //Load 4 components
                Load4Components(campaign);
                //Load station 5 assemble rru

                //Load remaining station

                svSftp.Disconnect();
            }
            else
            {
                Console.WriteLine("[Error] Unable to connect to sftp server");
            }

        }
        static void Load4Components(tblCampaign campaign)
        {
            var folders = new List<string>()
            {
                "1.1 Do kiem TRX",
                "1.2 Do kiem PA",
                "1.3 Do kiem Filter",
                "1.4 Do kiem Antenna"
            };
            string remoteDirectory = campaign.LogPath;
            string localDirectory = svSftp.LocalDirectory;

            if (svSftp.sftp.Exists(remoteDirectory) == false)
            {
                Console.WriteLine("[Warning] Folder \"{0}\" of campaign not exist.", remoteDirectory);
            }
            else
            {
                foreach (var folder in folders)
                {
                    if (svSftp.sftp.Exists(remoteDirectory + "/" + folder) == false)
                    {
                        Console.WriteLine("[Warning] Folder \"{0}\" not exist.", folder);
                        try
                        {
                            svSftp.sftp.CreateDirectory(remoteDirectory + "\\" + folder);
                            Console.WriteLine("[Warning] Created folder \"{0}\".", remoteDirectory + "\\" + folder);
                        }
                        catch (SftpPathNotFoundException) {
                            Console.WriteLine("[Error] Created folder failed: \"{0}\".", remoteDirectory + "\\" + folder);
                        }
                    }
                    else
                    {
                        Console.WriteLine("[Info] Reading log from folder: " + folder);
                        LoadFirst4Stations(remoteDirectory + "\\" + folder + "\\", localDirectory);
                    }
                }
            }
                
        }

        static void LoadFirst4Stations(string remoteDirectory, string localDirectory)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);
            foreach (SftpFile file in files)
            {
                if (!file.Name.StartsWith(".") && file.Name.Contains(".xlsx"))
                {
                    string localFilePath = Path.Combine(localDirectory, file.Name);
                    DownloadLogFile(localFilePath, file.FullName);

                    // If file found, read, save data to db, move remote file to storehouse then delete it
                    bool ret = ReadLogThenUpdateDb_1(localFilePath);
                    if (ret == true)
                        MoveFileToStorehouse(remoteDirectory, file);
                    else
                        Console.WriteLine("[Error] Update database failed!");
                    System.Threading.Thread.Sleep(100);
                    File.Delete(localFilePath);
                }  
            }
        }
        static bool ReadLogThenUpdateDb_1(string localfile)
        {
            bool ret = true;
            Excel exceltmp = new(localfile, 1);
            tblDevice newDevice = new(db, exceltmp);
            tblDeviceResult newDeviceResult = new(db, exceltmp, newDevice.id);

            if (db.InsertDevice(newDevice) == false ||
                db.InsertDeviceResult(newDeviceResult) == false ||
                InsertDetailResults(db, exceltmp, newDeviceResult.id) == false)
            {
                ret = false;
            }
            exceltmp.Close();
            return ret;
        }
        static void MoveFileToStorehouse(string remotePath, SftpFile file)
        {
            if (svSftp.sftp.Exists(remotePath + "storehouse") == false)
            {
                try
                {
                    svSftp.sftp.CreateDirectory(remotePath + "storehouse");
                }
                catch (SftpPathNotFoundException) { }
            }
            if (svSftp.sftp.Exists(remotePath + "storehouse/" + file.Name) == false)
            {
                file.MoveTo(remotePath + "storehouse/" + file.Name);
            }
            else
            {
                file.MoveTo(remotePath + "storehouse/" + DateTime.Now.ToString("MMddyyyy_HHmmss") + "_" + file.Name);
            }
        }

        static void DownloadLogFile(string localPath, string remotePath)
        {
            Stream stream;

            stream = File.OpenWrite(localPath);
            Console.WriteLine("[Info] Found file on sftp server: " + remotePath);
            Console.WriteLine("[Info] Downloading file to: " + localPath);
            svSftp.sftp.DownloadFile(remotePath, stream, x => Console.WriteLine("[Info] File's size: " + x));
            stream.Close();
        }

        static void LoadOneStation(string remoteDirectory, string localDirectory)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);

            foreach (var file in files)
            {
                string remoteFileName = file.Name;
                if (!file.Name.StartsWith(".") && file.Name.Contains(".xlsx"))
                {
                    Console.WriteLine("[Info] Found file on sftp server: " + file.FullName);
                    Stream stream = File.OpenWrite(localDirectory + file.Name);
                    string localFilePath = Path.Combine(localDirectory, file.Name);
                    svSftp.sftp.DownloadFile(file.FullName, stream, x => Console.WriteLine("[Info] File's size: " + x));
                    Console.WriteLine("[Info] Downloading file to: " + localFilePath);
                    stream.Close();
                    if (File.Exists(localFilePath))
                    {
                        // If file found, read, save data to db, move remote file to storehouse then delete it
                        if (ReadLogThenUpdateDb(localFilePath) == true)
                        {
                            if (svSftp.sftp.Exists(remoteDirectory + "storehouse") == false)
                            {
                                try
                                {
                                    svSftp.sftp.CreateDirectory(remoteDirectory + "storehouse");
                                }
                                catch (SftpPathNotFoundException) { }
                            }
                            if (svSftp.sftp.Exists(remoteDirectory + "storehouse/" + file.Name) == false)
                            {
                                file.MoveTo(remoteDirectory + "storehouse/" + file.Name);
                            }
                            else
                            {
                                file.MoveTo(remoteDirectory + "storehouse/" + DateTime.Now.ToString("MMddyyyy_HHmmss") + "_" + file.Name);
                            }
                        }
                        else
                        {
                            Console.WriteLine("[Error] Update database failed!");
                        }
                        System.Threading.Thread.Sleep(100);
                        File.Delete(localFilePath);
                        Console.WriteLine("[Info] Local File deleted.");
                    }
                }
                System.Threading.Thread.Sleep(100);
            }
        }
        static bool ReadLogThenUpdateDb(string localfile)
        {
            bool ret = true;
            Excel exceltmp = new(localfile, 1);

            tblDevice newDevice = new(db, exceltmp);
            tblDeviceResult newDeviceResult = new(db, exceltmp, newDevice.id);
            if (db.InsertDevice(newDevice) == false ||
                db.InsertDeviceResult(newDeviceResult) == false ||
                InsertDetailResults(db, exceltmp, newDeviceResult.id) == false)
            {
                ret = false;
            }
            exceltmp.Close();
            return ret;
        }
        static bool InsertDetailResults(ServerDatabase db, Excel excel, int newDeviceResultId)
        {
            bool ret = true;
            int lastRow = excel.ws.UsedRange.Rows.Count;
            using(var progress = new ProgressBar())
            {
                for (int excelRow = 11; excelRow <= lastRow; excelRow++)
                {
                    tblDetailResult newDetailResult = new(db, excel, excelRow, newDeviceResultId);
                    if (db.InsertDetailResult(newDetailResult) != false)
                    {
                        progress.Report((double)(excelRow - 11) / (lastRow - 11));
                    }
                    else
                    {
                        ret = false;
                        break;
                    }
                }
            }
            return ret;
        }

        static void RunCalibStationAPI()
        {
            var localIp = IPAddress.Any;
            var localPort = 1308;
            var localEndPoint = new IPEndPoint(localIp, localPort);
            var listener = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            listener.Bind(localEndPoint);
            listener.Listen(10);
            Console.WriteLine($"[Info] Local socket bind to {localEndPoint}. Waiting for request ...");

            var size = 1024;
            var receiveBuffer = new byte[size];
            while (true)
            {
                // tcp đòi hỏi một socket thứ hai làm nhiệm vụ gửi/nhận dữ liệu
                // socket này được tạo ra bởi lệnh Accept
                var socket = listener.Accept();
                Console.WriteLine($"[Info] Accepted connection from {socket.RemoteEndPoint}");

                // nhận dữ liệu vào buffer
                var length = socket.Receive(receiveBuffer);
                // không tiếp tục nhận dữ liệu nữa
                socket.Shutdown(SocketShutdown.Receive);
                var text = Encoding.ASCII.GetString(receiveBuffer, 0, length);
                Console.WriteLine($"[Info] Received: {text}");
                try
                {
                    JObject ob = JObject.Parse(text);
                    JToken sn_rru = ob["sn_rru"];
                    ComponentsSerialNumber myComponent = GetComponentSN(sn_rru.ToString());

                    var result = JsonConvert.SerializeObject(myComponent);
                    var sendBuffer = Encoding.ASCII.GetBytes(result);
                    // gửi kết quả lại cho client
                    socket.Send(sendBuffer);
                    Console.WriteLine($"[Info] Sent: {result}");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[Error] Exception: " + e.Message);
                    var sendBuffer = Encoding.ASCII.GetBytes(e.Message);
                    // gửi kết quả lại cho client
                    socket.Send(sendBuffer);
                }
                // không tiếp tục gửi dữ liệu nữa
                socket.Shutdown(SocketShutdown.Send);

                // đóng kết nối và giải phóng tài nguyên
                Console.WriteLine($"[Info] Closing connection from {socket.RemoteEndPoint}\r\n");
                socket.Close();
                Array.Clear(receiveBuffer, 0, size);
            }
        }
        static ComponentsSerialNumber GetComponentSN(string rru_sn)
        {
            ComponentsSerialNumber myCompnt = new ComponentsSerialNumber();

            try
            {
                Console.WriteLine("[Info] Openning Connection ...");
                string queryString = "select * from tbl_device where rru_sn = '" + rru_sn + "'";
                Console.WriteLine("[Info] queryString: " + queryString);
                using var command = new MySqlCommand(queryString, db.conn);
                MySqlDataReader rdr = command.ExecuteReader();

                while (rdr.Read())
                {
                    Console.WriteLine("[Info] " +rdr["id"] + "--" + rdr["mac"] + "--" + rdr["sn"] + "--" + rdr["rru_sn"]);
                    var sn_component = rdr["sn"].ToString();
                    if (sn_component.IndexOf("TRX") == 0)
                    {
                        myCompnt.sn_trx = sn_component;
                    }
                    else if (sn_component.IndexOf("PA") >= 0)
                    {
                        myCompnt.sn_pa = sn_component;
                    }
                    else if (sn_component.IndexOf("FIL") >= 0)
                    {
                        myCompnt.sn_fil = sn_component;
                    }
                    else if (sn_component.IndexOf("PWR") >= 0)
                    {
                        myCompnt.sn_pwr = sn_component;
                    }
                    else if (sn_component.IndexOf("ANT") >= 0)
                    {
                        myCompnt.sn_ant = sn_component;
                    }
                }
                rdr.Close();


                //var result = JsonConvert.SerializeObject(myCompnt);
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error] exception: " + e.Message);
            }
            return myCompnt;
        }
    }
}