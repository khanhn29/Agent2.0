using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Renci.SshNet.Common;

namespace Agent2._0
{
    class Program
    {
        static InfoServerDB dbSvInfo = new InfoServerDB();
        static InfoServerSftp sftpSvInfo = new InfoServerSftp();
        static ServerDatabase db;
        static ServerSftp svSftp;
        static void Main(string[] args)
        {
            Console.Title = "Agent 2.0";
            // Init server
            db = new ServerDatabase(dbSvInfo);
            svSftp = new ServerSftp(sftpSvInfo);
            if (db.conn.State == System.Data.ConnectionState.Open)
            {
                Console.WriteLine("Connected to Database");
                LoadLogToDB();
            }
            else
            {
                Console.WriteLine("Error: Unable to connect to database");
            }
            RunCalibStationAPI();
            Console.ReadLine();
        }

        static void LoadOneStation(string remoteDirectory, string localDirectory)
        {
            var files = svSftp.sftp.ListDirectory(remoteDirectory);

            foreach (var file in files)
            {
                string remoteFileName = file.Name;
                if (!file.Name.StartsWith(".") && file.Name.Contains(".xlsx"))
                {
                    Console.WriteLine("Sftp found file: " + file.FullName);
                    Stream stream = File.OpenWrite(localDirectory + file.Name);
                    string localFilePath = Path.Combine(localDirectory, file.Name);
                    //Download file
                    svSftp.sftp.DownloadFile(file.FullName, stream, x => Console.WriteLine(x));
                    Console.WriteLine("Downloading file to: " + localFilePath);
                    stream.Close();
                    // Check if file exists with its full path
                    if (File.Exists(localFilePath))
                    {
                        // If file found, read, save data to db ,move remote file to storehouse then delete it
                        if (ReadLog_UpdateDb(localFilePath) == true)
                        {
                            //Create folder
                            if (svSftp.sftp.Exists(remoteDirectory + "storehouse") == false)
                            {
                                try
                                {
                                    svSftp.sftp.CreateDirectory(remoteDirectory + "storehouse");
                                }
                                catch (SftpPathNotFoundException) { }
                            }
                            //Move remote file
                            if (svSftp.sftp.Exists(remoteDirectory + "storehouse/" + file.Name) == false){
                                file.MoveTo(remoteDirectory + "storehouse/" + file.Name);
                            }
                            else{
                                file.MoveTo(remoteDirectory + "storehouse/" + DateTime.Now.ToString("MMddyyyy_HHmmss") + "_" + file.Name);
                            }

                        }
                        else
                        {
                            Console.WriteLine("Update database failed!");
                        }
                        System.Threading.Thread.Sleep(100);
                        File.Delete(localFilePath);
                        Console.WriteLine("Local File deleted.");
                    }
                }
                System.Threading.Thread.Sleep(100);
            }
        }
        static void LoadLogToDB()
        {
            svSftp.Connect();
            if (svSftp.sftp.IsConnected)
            {
                string remoteDirectory = svSftp.RemoteDirectory;
                string localDirectory = svSftp.LocalDirectory;
                LoadOneStation(remoteDirectory, localDirectory);
                svSftp.Disconnect();
            }
            else
            {
                Console.WriteLine("Unable to connect to sftp server");
            }
        }

        static bool ReadLog_UpdateDb(string localfile)
        {
            bool ret = true;
            Excel exceltmp = new(localfile, 1);

            tblDevice newDevice = new(db, exceltmp);
            tblDeviceResult newDeviceResult = new(db, exceltmp, newDevice.id);
            if (db.InsertDevice(newDevice) == false ||
                db.InsertDeviceResult(newDeviceResult) == false ||
                ReadLog_UpdateDetailResult(db, exceltmp, newDeviceResult.id) == false)
            {
                ret = false;
            }
            exceltmp.Close();
            return ret;
        }
        static bool ReadLog_UpdateDetailResult(ServerDatabase db, Excel excel, int newDeviceResultId)
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
            Console.WriteLine($"Local socket bind to {localEndPoint}. Waiting for request ...");

            var size = 1024;
            var receiveBuffer = new byte[size];
            while (true)
            {
                // tcp đòi hỏi một socket thứ hai làm nhiệm vụ gửi/nhận dữ liệu
                // socket này được tạo ra bởi lệnh Accept
                var socket = listener.Accept();
                Console.WriteLine($"Accepted connection from {socket.RemoteEndPoint}");

                // nhận dữ liệu vào buffer
                var length = socket.Receive(receiveBuffer);
                // không tiếp tục nhận dữ liệu nữa
                socket.Shutdown(SocketShutdown.Receive);
                var text = Encoding.ASCII.GetString(receiveBuffer, 0, length);
                Console.WriteLine($"Received: {text}");
                try
                {
                    JObject ob = JObject.Parse(text);
                    JToken sn_rru = ob["sn_rru"];
                    ComponentsSerialNumber myComponent = GetComponentSN(sn_rru.ToString());

                    var result = JsonConvert.SerializeObject(myComponent);
                    var sendBuffer = Encoding.ASCII.GetBytes(result);
                    // gửi kết quả lại cho client
                    socket.Send(sendBuffer);
                    Console.WriteLine($"Sent: {result}");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: " + e.Message);
                    var sendBuffer = Encoding.ASCII.GetBytes(e.Message);
                    // gửi kết quả lại cho client
                    socket.Send(sendBuffer);
                }
                // không tiếp tục gửi dữ liệu nữa
                socket.Shutdown(SocketShutdown.Send);

                // đóng kết nối và giải phóng tài nguyên
                Console.WriteLine($"Closing connection from {socket.RemoteEndPoint}\r\n");
                socket.Close();
                Array.Clear(receiveBuffer, 0, size);
            }
        }
        static ComponentsSerialNumber GetComponentSN(string rru_sn)
        {
            ComponentsSerialNumber myCompnt = new ComponentsSerialNumber();

            try
            {
                Console.WriteLine("Openning Connection ...");
                string queryString = "select * from tbl_device where rru_sn = '" + rru_sn + "'";
                Console.WriteLine("queryString: " + queryString);
                using var command = new MySqlCommand(queryString, db.conn);
                MySqlDataReader rdr = command.ExecuteReader();

                while (rdr.Read())
                {
                    Console.WriteLine(rdr["id"] + "--" + rdr["mac"] + "--" + rdr["sn"] + "--" + rdr["rru_sn"]);
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
                var result = JsonConvert.SerializeObject(myCompnt);
                rdr.Close();

            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message);
            }
            return myCompnt;
        }
    }
}