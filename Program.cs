using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
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
        static List<string> folders = new List<string>()
        {
            "1.1 Do kiem TRX",
            "1.2 Do kiem PA",
            "1.3 Do kiem Filter",
            "1.4 Do kiem Antenna",
            "2. Lap rap RRU",
            "3.1 Luu Serial vao EEPROM RRU",
            "3.2 Do kiem Tx, Rx",
            "4. Test kin khi",
            "5. Burn in",
            "6. Do kiem TX RX sau burn in",
            "7. Test kin khi sau burn in",
            "8.1 OQC Chu trinh nhiet",
            "8.2 OQC Rung xoc",
            "9. Do kiem TX RX sau rung xoc, chu trinh nhiet",
            "10. Package"
        };
        static void Main(string[] args)
        {
            Console.Title = "Agent 2.0";
            // Init server
            db = new ServerDatabase(dbSvInfo);
            svSftp = new ServerSftp(sftpSvInfo);

            if (db.conn.State == System.Data.ConnectionState.Open)
            {
                Log.Info("Connected to Database");
                LoadCampaigns();
            }
            else
            {
                Log.Error("Unable to connect to database");
            }
            RunCalibStationAPI();
            Console.ReadLine();
        }

        static void LoadCampaigns()
        {
            List<tblCampaign> ActiveCampaignList = new List<tblCampaign>();
            MySqlDataReader rdr = db.Reader("SELECT id, name, from_date, to_date, log_path FROM tbl_campaign");
            while (rdr.Read())
            {
                DateTime fromdate = Convert.ToDateTime(rdr["from_date"]);
                DateTime todate = Convert.ToDateTime(rdr["to_date"]);
                Int32 id = rdr.GetInt32(0);
                string campaignName = rdr["name"].ToString();
                string logPath = rdr["log_path"].ToString();

                if (DateTime.Compare(fromdate, now) <= 0 && DateTime.Compare(now, todate) <= 0)
                {
                    Log.Info("Found campaign: " + id + "--" + campaignName + "--" + fromdate + "--" + todate + "--" + logPath);
                    ActiveCampaignList.Add(new tblCampaign(id, fromdate, todate, logPath));
                }
            }
            rdr.Close();

            foreach (var campaign in ActiveCampaignList)
            {
                LoadCampaign(campaign);
            }

        }

        static void LoadCampaign(tblCampaign campaign)
        {
            svSftp.Connect();
            if (svSftp.sftp.IsConnected)
            {
                ValidateCampaignfolder(campaign);

                LoadAllStations(campaign);
                svSftp.Disconnect();
            }
            else
            {
                Log.Error("Unable to connect to sftp server");
            }

        }
        static void ValidateCampaignfolder(tblCampaign campaign)
        {
            
            string remoteDirectory = campaign.LogPath;
            if (svSftp.Exists(remoteDirectory) == false)
            {
                Log.Warning("Folder \"{0}\" of campaign not exist.", remoteDirectory);
                svSftp.CreateDirectoryRecursively(remoteDirectory);
            }
            foreach (var folder in folders)
            {
                if (svSftp.Exists(remoteDirectory + "/" + folder) == false)
                {
                    Log.Warning("Folder \"{0}\" not exist.", folder);
                    try
                    {
                        svSftp.CreateDirectoryRecursively(remoteDirectory + "\\" + folder);
                        Log.Warning("Created folder \"{0}\".", remoteDirectory + "\\" + folder);
                    }
                    catch (SftpPathNotFoundException)
                    {
                        Log.Error("Created folder failed: \"{0}\".", remoteDirectory + "\\" + folder);
                    }
                }
            }
        }
        
        static void LoadAllStations(tblCampaign campaign)
        {
            string remotePath = campaign.LogPath;
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string localPath = System.IO.Path.GetDirectoryName(path);


            foreach (var folder in folders)
            {
                Log.Info("Reading log from folder: " + folder);
                LoadOneStation(remotePath + "\\" + folder + "\\", localPath, campaign);
            }
        }

        static void LoadOneStation(string remoteDirectory, string localDirectory, tblCampaign campaign)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);

            foreach (SftpFile remoteFile in files)
            {
                string remoteFileName = remoteFile.Name;
                if (!remoteFile.Name.StartsWith(".") && !remoteFile.Name.StartsWith("~") && remoteFile.Name.Contains(".xlsx"))
                {
                    Log.Info("Found file on sftp server: " + remoteFile.FullName);
                    if(VerifyFileSuccess(remoteFile, campaign))
                    {
                        string localFilePath = Path.Combine(localDirectory, remoteFile.Name);
                        DownloadFile(remoteFile.FullName, localFilePath);
                        if (File.Exists(localFilePath))
                        {
                            bool ret = UpdateFileToSQLDatabase(localFilePath);
                            if (ret == true)
                                MoveFileToStorehouse(remoteDirectory, remoteFile);
                            else
                                Log.Error("Update database failed!");
                        }
                        else
                        {
                            Log.Error("Download file failed");
                        }
                        System.Threading.Thread.Sleep(200);
                        File.Delete(localFilePath);
                    }
                    else
                    {
                        Log.Error("File not valid to insert to db");
                    }
                }
            }

        }
        static bool UpdateFileToSQLDatabase(string localfile)
        {
            bool ret = true;
            Excel exceltmp = new(localfile, 1);
            //tblDevice newDevice = new(db, exceltmp);
            //tblDeviceResult newDeviceResult = new(db, exceltmp, newDevice.id);

            exceltmp.Close();
            return ret;
        }
        static bool VerifyFileSuccess(SftpFile file, tblCampaign campaign)
        {
            bool ret = false;
            string FileName = file.Name;
            string[] parts = FileName.Split("_");
            DateTime fromDate = campaign.FromDate;
            DateTime toDate = campaign.ToDate;
            string sn = parts[0];

            try
            {
                Int32 nExist = db.Count("SELECT COUNT(id) FROM tbl_import_mac_sn WHERE sn='" + sn + "'");
                DateTime fileDate = DateTime.ParseExact(parts[1], "ddMMyyyy", CultureInfo.InvariantCulture);
                string msg = string.Format("File time info: fromDate[{0}] toDate[{1}] fileDate[{2}] macsnID[{3}]", fromDate, toDate, fileDate, nExist);
                Log.Info(msg);

                if (fromDate <= fileDate && fileDate <= toDate && nExist > 0)
                {
                    ret = true;
                }
            }
            catch(Exception e)
            {
                Log.Error("File name incorrect format" + e.Message);
            }

            return ret;
        }
        static void DownloadFile(string remotePath, string localPath)
        {
            try
            {
                Stream stream;

                stream = File.OpenWrite(localPath);
                Log.Info("Found file on sftp server: " + remotePath);
                Log.Info("Downloading file to: " + localPath);
                svSftp.sftp.DownloadFile(remotePath, stream, x => Log.Info("File's size: " + x));
                stream.Close();
            }
            catch (Exception e)
            {
                Log.Error("" + e.Message);
                KillSpecificExcelFileProcess();
            }
        }
        static void MoveFileToStorehouse(string remotePath, SftpFile file)
        {
            if (svSftp.Exists(remotePath + "\\storehouse") == false)
            {
                try
                {
                    svSftp.sftp.CreateDirectory(remotePath + "\\storehouse");
                }
                catch (SftpPathNotFoundException) { }
            }
            if (svSftp.sftp.Exists(remotePath + "\\storehouse\\" + file.Name) == false)
            {
                file.MoveTo(remotePath + "\\storehouse\\" + file.Name);
            }
            else
            {
                file.MoveTo(remotePath + "\\storehouse\\" + DateTime.Now.ToString("MMddyyyy_HHmmss") + "_" + file.Name);
            }
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
                    Log.Error("Exception: " + e.Message);
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
            myCompnt.sn_rru = rru_sn;
            try
            {
                Log.Info("Openning Connection ...");
                string queryString = "select * from tbl_device where rru_sn = '" + rru_sn + "'";
                Log.Info("queryString: " + queryString);
                using var command = new MySqlCommand(queryString, db.conn);
                MySqlDataReader rdr = command.ExecuteReader();

                while (rdr.Read())
                {
                    Log.Info("" + rdr["id"] + "--" + rdr["mac"] + "--" + rdr["sn"] + "--" + rdr["rru_sn"]);
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
                Log.Error("exception: " + e.Message);
            }
            return myCompnt;
        }
        private static void KillSpecificExcelFileProcess()
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel")
                    process.Kill();
            }
        }
    }
}