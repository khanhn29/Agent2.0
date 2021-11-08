﻿using System;
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
        static List<string> StationsNameList = new List<string>
        {
            "TRX-TEST",
            "PA-TEST",
            "FILTER-TEST",
            "ANT-TEST",
            "ASSEM-RRU"
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
                Log.Info("");
            }
        }

        static void LoadOneStation(string remoteDirectory, string localDirectory, tblCampaign campaign)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);

            foreach (SftpFile remoteFile in files)
            {
                if (!remoteFile.Name.StartsWith(".") && !remoteFile.Name.StartsWith("~") && remoteFile.Name.Contains(".xlsx"))
                {
                    Log.Info("=============Start=============");
                    Log.Info("Found file on sftp server: " + remoteFile.FullName);
                    if(IsFileValid(remoteFile, campaign)){
                        string localFilePath = Path.Combine(localDirectory, remoteFile.Name);
                        DownloadFile(remoteFile.FullName, localFilePath);
                        if (File.Exists(localFilePath)){
                            bool ret = UpdateFileToSQLDatabase(localFilePath);
                            if (ret == true)
                                MoveFileToStorehouse(remoteDirectory, remoteFile);
                            else
                                Log.Error("Update database failed!");
                        }
                        else{
                            Log.Error("Download file failed");
                        }
                        System.Threading.Thread.Sleep(100);
                        File.Delete(localFilePath);
                    }
                    else{
                        Log.Error("File not valid to insert to db");
                    }
                    Log.Info("=============End=============");
                }
            }

        }
        static bool UpdateFileToSQLDatabase(string localfile)
        {
            bool ret = true;

            Excel exceltmp = new(localfile, 1);
            var match = StationsNameList
                .FirstOrDefault(stringToCheck => stringToCheck.Contains(exceltmp.StationName));

            if (match != null)
            {
                tblDevice newDevice = new(db, exceltmp);
                db.DeactivateOldDevices(newDevice);
                ret = db.InsertDevice(newDevice);
                if (ret == true)
                {
                    string msg = string.Format("Inserted device Id[{0}] Mac[{1}] Mac2[{2}] SN[{3}]",
                        newDevice.id, newDevice.mac, newDevice.mac2, newDevice.sn);
                    Log.Info(msg);
                }
                if(exceltmp.FileSerialNum.StartsWith("RRU"))
                {
                    ComponentsSerialNumber componentsInfo = GetComponentSNInExcel(db, exceltmp);
                    FillRRUSN2Components(componentsInfo);
                }

            }

            //tblDeviceResult newDeviceResult = new(db, exceltmp, newDevice.id);

            exceltmp.Close();
            return ret;
        }
        static bool IsFileValid(SftpFile file, tblCampaign campaign)
        {
            bool ret = true;
            string FileName = file.Name;
            DateTime fromDate = campaign.FromDate;
            DateTime toDate = campaign.ToDate;
            string[] parts = FileName.Split("_");
            string sn = parts[0];
            string stationName = parts[1];
            string timeHHmm = parts[2];
            string dateddMMyyyy = parts[3];
            string result = parts[4];

            try
            {
                Int32 nExist = db.Count("SELECT COUNT(id) FROM tbl_import_mac_sn WHERE sn='" + sn + "'");
                DateTime fileDate = DateTime.ParseExact(dateddMMyyyy, "ddMMyyyy", CultureInfo.InvariantCulture);
                string msg = string.Format("File time info: fromDate[{0}] fileDate[{2}] toDate[{1}] macsnID[{3}]", fromDate, toDate, fileDate, nExist);
                Log.Info(msg);

                if(fileDate < fromDate || toDate < fileDate)
                {
                    Log.Error("File " + file.Name + " is not in campaign's duration");
                    ret = false;
                }
                if(nExist == 0)
                {
                    Log.Error("SN " + sn + " is not in campaign's plan");
                    ret = false;
                }
            }
            catch(Exception e)
            {
                Log.Error("File name incorrect format" + e.Message);
                ret = false;
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
                    ComponentsSerialNumber myComponent = GetComponentSNInSQLDatabase(sn_rru.ToString());

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
        private static bool FillRRUSN2Components(ComponentsSerialNumber info)
        {
            bool ret = false;
            if (info.sn_trx == "")
            {
                Log.Error("Upate RRU SN to components failed: " + info.sn_trx + ": " + info.mac + "--" + info.mac2 + "--" + info.sn_rru);
            }
            else
            {
                try
                {
                    string insertQuery = "UPDATE tbl_device SET mac=@mac, mac2=@mac2, rru_sn=@rru_sn WHERE sn=@trx_sn";
                    MySqlCommand cmd = new MySqlCommand(insertQuery, db.conn);
                    cmd.Parameters.Add("?mac", MySqlDbType.String).Value = info.mac;
                    cmd.Parameters.Add("?mac2", MySqlDbType.String).Value = info.mac2;
                    cmd.Parameters.Add("?rru_sn", MySqlDbType.String).Value = info.sn_rru;
                    cmd.Parameters.Add("?trx_sn", MySqlDbType.String).Value = info.sn_trx;
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        Log.Info("Update Device: " + info.sn_trx + ": " + info.mac + "--" + info.mac2 + "--" + info.sn_rru);
                        ret = true;
                    }
                    else
                    {
                        Log.Error("FillRRUSN2Components failed: " + info.sn_trx + ": " + info.mac + "--" + info.mac2 + "--" + info.sn_rru);
                        ret = false;
                    }

                    insertQuery = "UPDATE tbl_device SET rru_sn=@rru_sn WHERE sn=@pa_sn";
                    cmd = new MySqlCommand(insertQuery, db.conn);
                    cmd.Parameters.Add("?rru_sn", MySqlDbType.String).Value = info.sn_rru;
                    cmd.Parameters.Add("?pa_sn", MySqlDbType.String).Value = info.sn_pa;
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        Log.Info("Update Device: " + info.sn_pa + ": " + info.sn_rru);
                        ret = true;
                    }
                    else
                    {
                        Log.Error("FillRRUSN2Components failed: " + info.sn_pa + ": " + info.sn_rru);
                        ret = false;
                    }

                    insertQuery = "UPDATE tbl_device SET rru_sn=@sn_rru WHERE sn=@sn_fil";
                    cmd = new MySqlCommand(insertQuery, db.conn);
                    cmd.Parameters.Add("?sn_rru", MySqlDbType.String).Value = info.sn_rru;
                    cmd.Parameters.Add("?sn_fil", MySqlDbType.String).Value = info.sn_fil;
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        Log.Info("Update Device: " + info.sn_fil + ": " + info.sn_rru);
                        ret = true;
                    }
                    else
                    {
                        Log.Error("Upate RRU SN to components failed: " + info.sn_fil + ": " + info.sn_rru);
                        ret = false;
                    }

                    insertQuery = "UPDATE tbl_device SET rru_sn=@sn_rru WHERE sn=@sn_ant";
                    cmd = new MySqlCommand(insertQuery, db.conn);
                    cmd.Parameters.Add("?sn_rru", MySqlDbType.String).Value = info.sn_rru;
                    cmd.Parameters.Add("?sn_ant", MySqlDbType.String).Value = info.sn_ant;
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        Log.Info("Update Device: " + info.sn_ant + ": " + info.sn_rru);
                        ret = true;
                    }
                    else
                    {
                        Log.Error("Upate RRU SN to components failed: " + info.sn_ant + ": " + info.sn_rru);
                        ret = false;
                    }
                }
                catch (MySqlException ex)
                {
                    Log.Error("Insert Device: " + ex.Message);
                    ret = false;
                }

            }

            return ret;
        }
        private static ComponentsSerialNumber GetComponentSNInExcel(ServerDatabase db, Excel excel)
        {
            ComponentsSerialNumber ret = new();

            ret.sn_rru = excel.ReadCell(6, 2);
            if (!ret.sn_rru.StartsWith("RRU"))
            {
                Log.Error("RRU SN in assemble station do not satisfy the format RRU******");
                ret.sn_rru = "";
            }

            ret.sn_trx = excel.ReadCell(13, 4);
            if (!ret.sn_trx.StartsWith("TRX"))
            {
                Log.Error("TRX SN in assemble station do not satisfy the format TRX******");
                ret.sn_trx = "";
            }
            else
            {
                string macAddress = excel.ReadCell(7, 2);
                string[] macs = macAddress.Split("/");
                ret.mac = macs[0];
                ret.mac2 = macs[1];
            }

            ret.sn_pa = excel.ReadCell(14, 4);
            if (!ret.sn_pa.StartsWith("PA"))
            {
                Log.Error("PA SN in assemble station do not satisfy the format PA******");
                ret.sn_pa = "";
            }

            ret.sn_fil = excel.ReadCell(15, 4);
            if (!ret.sn_fil.StartsWith("FILTER"))
            {
                Log.Error("FILTER SN in assemble station do not satisfy the format FILTER******");
                ret.sn_fil = "";
            }

            ret.sn_ant = excel.ReadCell(16, 4);
            if (!ret.sn_ant.StartsWith("ANT"))
            {
                Log.Error("ANT SN in assemble station do not satisfy the format ANT******");
                ret.sn_ant = "";
            }

            return ret;
        }
        static ComponentsSerialNumber GetComponentSNInSQLDatabase(string rru_sn)
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