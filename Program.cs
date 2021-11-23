﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Renci.SshNet.Common;
using Renci.SshNet.Sftp;
using IExcel = Microsoft.Office.Interop.Excel;

namespace Agent2._0
{
    class Program
    {
        public static bool DebugLogOn = false;
        static ServerDatabase db;
        static ServerSftp svSftp;
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
            "ASSEM-RRU",
            "SAVE-DATA",
            "PERFORMANCE-TEST",
            "AIR-TEST",
            "RRU-BURN-IN",
            "TEST-TRX-BURN-IN",
            "AIR-TEST-AFTER-BURN-IN",
            "THERMAL-CYCLE",
            "VIBRATION-TEST",
            "PERFORMANCE-TEST-AFTER-VIBRATION",
            "PACKAGE"
        };
        static void Main(string[] args)
        {
#pragma warning disable CA1416 // Validate platform compatibility
            Console.SetWindowSize(180, 40);
#pragma warning restore CA1416 // Validate platform compatibility
            Console.Title = "Agent 2.0";

            DisableConsoleQuickEdit.Go();
            try
            {
                bool successfulConnectionSFTP = false;
                while(!successfulConnectionSFTP)
                {
                    System.Threading.Thread.Sleep(1000);
                    successfulConnectionSFTP = ConnectSFTP();
                }
                bool successfulConnectionMySQP = false;
                while (!successfulConnectionMySQP)
                {
                    System.Threading.Thread.Sleep(1000);
                    successfulConnectionMySQP = ConnectMySQL();
                }
                KillSpecificExcelFileProcess();
                Thread StationAPIThrd = new(RunCalibStationAPI);
                Thread LoadLogThrd = new(LoadCampaigns);

                StationAPIThrd.Start();
                LoadLogThrd.Start();
            }
            catch(Exception e)
            {
                Log.Error(e.Message);
            }
            //svSftp.Disconnect();
            Console.ReadLine();
        }

        static bool ConnectSFTP()
        {
            InfoServerSftp sftpSvInfo = new InfoServerSftp();
            svSftp = new ServerSftp(sftpSvInfo);
            svSftp.Connect();
            return svSftp.sftp.IsConnected;
        }
        static bool ConnectMySQL()
        {
            InfoServerDB dbSvInfo = new InfoServerDB();
            db = new ServerDatabase(dbSvInfo);
            db.Open();
            return db.conn.State == System.Data.ConnectionState.Open;
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
                DateTime now = DateTime.Now;
                if (DateTime.Compare(fromdate, now) <= 0 && DateTime.Compare(now, todate) <= 0)
                {
                    Log.Debug("Found campaign: " + id + "--" + campaignName + "--" + fromdate + "--" + todate + "--" + logPath);
                    ActiveCampaignList.Add(new tblCampaign(id, fromdate, todate, logPath));
                }
            }
            rdr.Close();

            foreach (var campaign in ActiveCampaignList)
            {
                LoadCampaign(campaign);
            }

            while(true)
            {
                System.Threading.Thread.Sleep(5000);
                foreach (var campaign in ActiveCampaignList)
                {
                    LoadCampaign(campaign);
                }
            }

        }

        static void UpdateNewFileInCampaign(tblCampaign campaign)
        {
            string remoteFolder = campaign.LogPath;
            foreach (string folder in folders)
            {
                string remoteStationFolder = remoteFolder + "/" + folder;
                var files = svSftp.sftp.ListDirectory(remoteStationFolder);
                foreach (SftpFile remoteFile in files)
                {
                    if (!remoteFile.Name.StartsWith(".") && !remoteFile.Name.StartsWith("~") && remoteFile.Name.Contains(".xlsx"))
                    {
                        Log.Info("Found new file: " + remoteFile.FullName);
                        if(IsFileValid(remoteFile, campaign))
                        {

                        }
                    }
                }
            }
        }

        static void LoadCampaign(tblCampaign campaign)
        {
            ValidateCampaignfolder(campaign);
            LoadAllStations(campaign);
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
                Log.Debug("Reading log from folder: " + folder);
                LoadOneStation(remotePath + "\\" + folder + "\\", localPath, campaign);
                Log.Debug("");
            }
        }

        static void LoadOneStation(string remoteDirectory, string localDirectory, tblCampaign campaign)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);

            foreach (SftpFile remoteFile in files)
            {
                if (!remoteFile.Name.StartsWith(".") && !remoteFile.Name.StartsWith("~") && remoteFile.Name.Contains(".xlsx"))
                {
                    Log.Debug("=============Start=============");
                    Log.Debug("Found file on sftp server: " + remoteFile.FullName);
                    if(IsFileValid(remoteFile, campaign))
                    {
                        string localFilePath = Path.Combine(localDirectory, remoteFile.Name);
                        DownloadFile(remoteFile.FullName, localFilePath);
                        if (File.Exists(localFilePath))
                        {
                            bool ret = UpdateFileToSQLDatabase(localFilePath, campaign);
                            if (ret == true)
                            {
                                MoveFileToStorehouse(remoteDirectory, remoteFile);
                                Log.Info("Update file to SQL success: " + remoteFile.FullName);
                            }
                            else
                            {
                                Log.Error("Update database failed!");
                            }
                        }
                        else{
                            Log.Error("Download file failed");
                        }
                        System.Threading.Thread.Sleep(100);
                        try
                        {
                            File.Delete(localFilePath);
                        }
                        catch(Exception e)
                        {
                            Log.Error(e.Message);
                        }
                    }
                    else{
                        Log.Error("Update file to SQL failed: " + remoteFile.FullName);
                        MoveFileToNotInserted(remoteDirectory, remoteFile);
                    }
                    Log.Debug("=============End=============");
                }
            }

        }
        static bool UpdateFileToSQLDatabase(string localfile, tblCampaign campaign)
        {
            bool ret = true;

            Excel exceltmp = new(localfile, 1);
            var match = StationsNameList.FirstOrDefault(stringToCheck => stringToCheck.Contains(exceltmp.StationName));

            if (match != null)
            {
                int countDV = db.Count("SELECT COUNT(id) from tbl_device WHERE sn='" + exceltmp.FileSerialNum + "'");
                tblDevice newDevice = new(db, exceltmp);

                if (countDV == 0)
                {
                    ret = db.InsertDevice(newDevice);
                    if (ret == true)
                    {
                        string msg = string.Format("Inserted device Id[{0}] Mac[{1}] Mac2[{2}] SN[{3}]",
                            newDevice.id, newDevice.mac, newDevice.mac2, newDevice.sn);
                        Log.Debug(msg);
                    }
                    else
                    {
                        Log.Error("Insert device failed");
                        goto Finish;
                    }
                }
                else
                {
                    Log.Warning("Device " + exceltmp.FileSerialNum + " exists in Database, will be update if needed");
                    db.UpdateDevice(newDevice);
                }

                if (exceltmp.StationName.Contains("ASSEM-RRU"))
                {
                    ComponentsSerialNumber componentsInfo = GetComponentSNInExcel(db, exceltmp);
                    componentsInfo.Print();
                    Log.Debug("Fill RRU Serial number");
                    //FillRRUSN2Components(componentsInfo);
                    FillRRUSerialNumToComponents_2(componentsInfo);
                }
            }

            tblDeviceResult newDeviceResult = new(db, exceltmp, campaign);
            db.ExecuteNonQuery("UPDATE tbl_device_result SET latest='0' where device_id='"+
                newDeviceResult.device_id + "' and station_name='" + newDeviceResult.station_name+ "'");
            ret = db.InsertDeviceResult(newDeviceResult);
            if (ret == true)
            {
                string msg = string.Format("Inserted DeviceResult Id[{0}] dvID[{1}] CampID[{2}]",
                    newDeviceResult.id, newDeviceResult.device_id, campaign.Id);
                Log.Debug(msg);
            }
            else
            {
                Log.Error("Insert device result failed");
                goto Finish;
            }

            ret = InsertDetailResults(db, exceltmp, newDeviceResult.id);
            if (ret == true)
            {
                Log.Debug("Insert detail result successful");
            }
            else
            {
                Log.Error("Insert detail result failed");
                goto Finish;
            }

        Finish:
            exceltmp.Close();
            return ret;
        }
        static bool InsertDetailResults(ServerDatabase db, Excel excel, int newDeviceResultId)
        {
            bool ret = true;
            long fullRow = excel.ws.UsedRange.Rows.Count;
            long lastRow = excel.ws.Cells[fullRow + 100, 1].End(IExcel.XlDirection.xlUp).Row;
            using (var progress = new ProgressBar())
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
        static bool IsFileValid(SftpFile file, tblCampaign campaign)
        {
            bool ret = true;
            string FileName = file.Name.Split(".")[0];
            DateTime fromDate = campaign.FromDate;
            DateTime toDate = campaign.ToDate;
            string[] parts = FileName.Split("_");
            if (parts.Length != 5)
            {
                Log.Error("File name is not match format <SerialNum>_<StationName>_<Time>_<Date>_<Result>");
                ret = false;
                goto finish;
            }
                
            string sn = parts[0];
            string stationName = parts[1];
            string timeHHmm = parts[2];
            string dateddMMyyyy = parts[3];
            string result = parts[4];

            try
            {
                DateTime fileDate = DateTime.ParseExact(dateddMMyyyy, "ddMMyyyy", CultureInfo.InvariantCulture);
                if (fileDate < fromDate || toDate < fileDate)
                {
                    Log.Error("File " + file.Name + " is not in campaign's duration");
                    ret = false;
                }
            }
            catch
            {
                Log.Error("Invalid Date format(ddMMyyyy) in file name: " + dateddMMyyyy);
                ret = false;
            }

            Int32 nExist = db.Count("SELECT COUNT(id) FROM tbl_import_mac_sn WHERE sn='" + sn + "'");
            if(nExist == 0)
            {
                Log.Error("SN " + sn + " is not in campaign's plan");
                ret = false;
            }

            var match = StationsNameList.FirstOrDefault(stringToCheck => stringToCheck.Contains(stationName));
            if (match == null)
            {
                Log.Error("Invalid station name: " + stationName);
                ret = false;
            }

            try
            {
                DateTime FileTime = DateTime.ParseExact(timeHHmm, "HHmm", CultureInfo.InvariantCulture);
            }
            catch
            {
                Log.Error("Invalid time format(HHmm) in file name: " + timeHHmm);
                ret = false;
            }

            if(result!="PASS" && result!="FAIL")
            {
                Log.Error("Invalid file name format result: " + result);
                ret = false;
            }

            finish:
            return ret;
        }
        static void DownloadFile(string remotePath, string localPath)
        {
            try
            {
                Stream stream;

                stream = File.OpenWrite(localPath);
                Log.Debug("Found file on sftp server: " + remotePath);
                Log.Debug("Downloading file to: " + localPath);
                svSftp.sftp.DownloadFile(remotePath, stream, x => Log.Debug("File's size: " + x));
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
        static void MoveFileToNotInserted(string remotePath, SftpFile file)
        {
            if (svSftp.Exists(remotePath + "\\NotInserted") == false)
            {
                try
                {
                    svSftp.sftp.CreateDirectory(remotePath + "\\NotInserted");
                }
                catch (SftpPathNotFoundException) { }
            }
            if (svSftp.sftp.Exists(remotePath + "\\NotInserted\\" + file.Name) == false)
            {
                file.MoveTo(remotePath + "\\NotInserted\\" + file.Name);
            }
            else
            {
                file.MoveTo(remotePath + "\\NotInserted\\" + DateTime.Now.ToString("MMddyyyy_HHmmss") + "_" + file.Name);
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
            Log.Info($"Local socket bind to {localEndPoint}. Waiting for request ...");

            var size = 1024;
            var receiveBuffer = new byte[size];
            while (true)
            {
                // tcp đòi hỏi một socket thứ hai làm nhiệm vụ gửi/nhận dữ liệu
                // socket này được tạo ra bởi lệnh Accept
                var socket = listener.Accept();
                Log.Info($"Accepted connection from {socket.RemoteEndPoint}");

                // nhận dữ liệu vào buffer
                var length = socket.Receive(receiveBuffer);
                // không tiếp tục nhận dữ liệu nữa
                socket.Shutdown(SocketShutdown.Receive);
                var text = Encoding.ASCII.GetString(receiveBuffer, 0, length);
                Log.Info($"Received: {text}");
                try
                {
                    JObject ob = JObject.Parse(text);
                    JToken sn_rru = ob["sn_rru"];
                    ComponentsSerialNumber myComponent = GetComponentSNInSQLDatabase(sn_rru.ToString());

                    var result = JsonConvert.SerializeObject(myComponent);
                    var sendBuffer = Encoding.ASCII.GetBytes(result);
                    // gửi kết quả lại cho client
                    socket.Send(sendBuffer);
                    Log.Info($"Sent: {result}");
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
                Log.Debug($"Closing connection from {socket.RemoteEndPoint}\r\n");
                socket.Close();
                Array.Clear(receiveBuffer, 0, size);
            }
        }
        private static bool FillRRUSN2Components(ComponentsSerialNumber info)
        {
            bool ret = true;
            string queryStr;
            MySqlCommand cmd;
            if (info.sn_trx == "")
            {
                string errMsg = string.Format("Fill SN failed: sn_rru[{0}], sn_trx[{1}], sn_pa[{2}], sn_fil[{3}], sn_ant[{4}], mac[{5}], mac2[{6}]",
                    info.sn_rru, info.sn_trx, info.sn_pa1, info.sn_fil, info.sn_ant, info.mac, info.mac2);
                Log.Error(errMsg);
            }
            else
            {
                try
                {
                    string rru_sn = db.GetString("SELECT rru_sn FROM tbl_device WHERE sn='" + info.sn_trx + "'");
                    if(rru_sn.CompareTo(info.sn_rru) != 0)
                    {
                        string history = db.GetString("SELECT history FROM tbl_device WHERE sn='" + info.sn_trx + "'");
                        history += string.Format("Old: {0}--New: {1}--Time:{2}--LogFile:{3}\r\n",
                            rru_sn, info.sn_rru, DateTime.Now.ToString(), info.FileName);
                        queryStr = string.Format("UPDATE tbl_device SET mac='{0}', mac2='{1}', rru_sn='{2}', history='{3}' WHERE sn='{4}'",
                            info.mac, info.mac2, info.sn_rru, history, info.sn_trx);
                        cmd = new MySqlCommand(queryStr, db.conn);
                        if (cmd.ExecuteNonQuery() == 1)
                        {
                            Log.Debug("Update Device: " + info.sn_trx + ": " + info.mac + "--" + info.mac2 + "--" + info.sn_rru);
                        }
                        else
                        {
                            Log.Error("Fill SN failed: " + info.sn_trx + ": " + info.mac + "--" + info.mac2 + "--" + info.sn_rru);
                            ret = false;
                        }
                    }

                    string[] listCMPNT =
                    {
                        info.sn_pa1,
                        info.sn_pa2,
                        info.sn_fil,
                        info.sn_ant
                    };

                    foreach(string cmpnt in listCMPNT)
                    {
                        rru_sn = db.GetString("SELECT rru_sn FROM tbl_device WHERE sn='" + cmpnt + "'");
                        if (rru_sn.CompareTo(info.sn_rru) != 0)
                        {
                            string history = db.GetString("SELECT history FROM tbl_device WHERE sn='" + cmpnt + "'");
                            history += string.Format("Old: {0}--New: {1}--Time:{2}--LogFile:{3}\r\n",
                                            rru_sn, info.sn_rru, DateTime.Now.ToString(), info.FileName);
                            queryStr = string.Format("UPDATE tbl_device SET mac='{0}', mac2='{1}', rru_sn='{2}', history='{3}' WHERE sn='{4}'",
                                            "", "", info.sn_rru, history, cmpnt);
                            cmd = new MySqlCommand(queryStr, db.conn);
                            if (cmd.ExecuteNonQuery() == 1)
                            {
                                Log.Debug("Update Device: " + cmpnt + ":" + info.sn_rru);
                            }
                            else
                            {
                                Log.Error("Fill SN failed: " + cmpnt + ":" + info.sn_rru);
                                ret = false;
                            }
                        }
                    }

                    queryStr = string.Format("UPDATE tbl_device SET mac='{0}', mac2='{1}', history='InsertTime:{2}' WHERE sn='{3}'",
                            info.mac, info.mac2, DateTime.Now.ToString(), info.sn_rru);
                    cmd = new MySqlCommand(queryStr, db.conn);
                    if (cmd.ExecuteNonQuery() == 1)
                    {
                        Log.Debug("Update Device: " + info.sn_rru);
                    }
                    else
                    {
                        Log.Error("Fill SN failed: " + info.sn_rru);
                        ret = false;
                    }
                }
                catch (MySqlException ex)
                {
                    Log.Error(ex.Message);
                    ret = false;
                }

            }

            return ret;
        }
        private static bool FillRRUSerialNumToComponents_2(ComponentsSerialNumber info)
        {
            bool ret = true;

            string[] listCMPNT =
                {
                    info.sn_rru,
                    info.sn_trx,
                    info.sn_fil,
                    info.sn_ant,
                    info.sn_pa1,
                    info.sn_pa2
                };
            foreach (string cmpnt in listCMPNT)
            {
                if (cmpnt == "")
                    continue;
                if(db.Count("SELECT COUNT(id) FROM tbl_device WHERE sn='" + cmpnt + "'") > 0)
                {
                    string currentRRUSN = db.GetString("SELECT rru_sn FROM tbl_device WHERE sn = '" + cmpnt + "'");
                    if (currentRRUSN != info.sn_rru)
                    {
                        string history = db.GetString("SELECT history FROM tbl_device WHERE sn='" + cmpnt + "'");
                        history += string.Format("Old: {0}--New: {1}--Time:{2}--LogFile:{3}\r\n",
                            currentRRUSN, info.sn_rru, DateTime.Now.ToString(), info.FileName);
                        string queryStr = "UPDATE tbl_device SET rru_sn='" + info.sn_rru
                            + "', history='" + history
                            + "' WHERE sn='"+ cmpnt + "'";
                        db.ExecuteNonQuery(queryStr);
                    }
                }
                else
                {
                    string history = string.Format("Insert time:" + DateTime.Now.ToString());
                    int id = db.GetInt16("SELECT MAX(id) FROM tbl_device") + 1;
                    string queryStr = string.Format("INSERT INTO tbl_device (id, mac, mac2, sn, rru_sn, history)" +
                        " VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}')",
                        id, "", "", cmpnt, info.sn_rru, history);
                    db.ExecuteNonQuery(queryStr);
                }
            }

            if(info.sn_trx != "")
            {
                string queryStr = string.Format("UPDATE tbl_device SET mac='{0}', mac2='{1}', rru_sn='{2}' WHERE sn='{3}'",
                                info.mac, info.mac2, info.sn_rru, info.sn_trx);
                db.ExecuteNonQuery(queryStr);
            }
            if (info.sn_rru != "")
            {
                string queryStr = string.Format("UPDATE tbl_device SET mac='{0}', mac2='{1}', rru_sn='{2}' WHERE sn='{3}'",
                                info.mac, info.mac2, info.sn_rru, info.sn_rru);
                db.ExecuteNonQuery(queryStr);
            }

            return ret;
        }
        private static ComponentsSerialNumber GetComponentSNInExcel(ServerDatabase db, Excel excel)
        {
            ComponentsSerialNumber ret = new();

            ret.FileName = excel.FileName;
            ret.sn_rru = excel.FileSerialNum;
            ret.sn_trx = excel.ReadCell(13, 4);
            ret.mac = db.GetString("SELECT mac FROM tbl_import_mac_sn WHERE sn = '" + ret.sn_trx + "'");
            ret.mac2 = db.GetString("SELECT mac2 FROM tbl_import_mac_sn WHERE sn = '" + ret.sn_trx + "'");
            ret.sn_fil = excel.ReadCell(14, 4);
            ret.sn_ant = excel.ReadCell(15, 4);
            ret.sn_pa1 = excel.ReadCell(16, 4);
            ret.sn_pa2 = excel.ReadCell(17, 4);

            return ret;
        }
        private static ComponentsSerialNumber GetComponentSNInSQLDatabase(string rru_sn)
        {
            ComponentsSerialNumber myCompnt = new ComponentsSerialNumber();
            myCompnt.sn_rru = rru_sn;
            try
            {
                Log.Debug("Openning Connection ...");
                string queryString = "select * from tbl_device where rru_sn = '" + rru_sn + "'";
                Log.Debug("queryString: " + queryString);
                using var command = new MySqlCommand(queryString, db.conn);
                MySqlDataReader rdr = command.ExecuteReader();

                while (rdr.Read())
                {
                    Log.Debug("" + rdr["id"] + "--" + rdr["mac"] + "--" + rdr["sn"] + "--" + rdr["rru_sn"]);
                    var sn_component = rdr["sn"].ToString();
                    if (sn_component.IndexOf("MTR") >= 0)
                    {
                        myCompnt.sn_trx = sn_component;
                    }
                    else if (sn_component.IndexOf("MPA") >= 0)
                    {
                        if (myCompnt.sn_pa1 != "")
                            myCompnt.sn_pa2 = sn_component;
                        else
                            myCompnt.sn_pa1 = sn_component;
                    }
                    else if (sn_component.IndexOf("MFL") >= 0)
                    {
                        myCompnt.sn_fil = sn_component;
                    }
                    else if (sn_component.IndexOf("ANT") >= 0)
                    {
                        myCompnt.sn_ant = sn_component;
                    }
                }
                rdr.Close();

                myCompnt.mac = db.GetString("Select mac from tbl_device where sn = '" + rru_sn + "' LIMIT 1");
                myCompnt.mac2 = db.GetString("Select mac2 from tbl_device where sn = '" + rru_sn + "' LIMIT 1");


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