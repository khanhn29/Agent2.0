using Renci.SshNet;
using System;

namespace Agent2._0
{
    class ServerSftp
    {
        public SftpClient sftp { get; set; }
        public string RemoteDirectory { get; set; }
        public string LocalDirectory { get; set; }
        public ServerSftp(InfoServerSftp sftpSvInfo)
        {
            string strExeFilePath = System.Reflection.Assembly.GetEntryAssembly().Location;
            string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);

            sftp = new SftpClient(new PasswordConnectionInfo(sftpSvInfo.Ip, sftpSvInfo.Username, sftpSvInfo.Pass));
            RemoteDirectory = GetRemoteDirectory();
            LocalDirectory = strWorkPath + "\\";
        }
        private string GetRemoteDirectory()
        {
            return @"\sftp\";
        }
        public void Connect()
        {
            try
            {
                sftp.ConnectionInfo.Timeout = TimeSpan.FromSeconds(10);
                sftp.Connect();
            }
            catch(Exception e)
            {
                Console.WriteLine("Connect sftp error: " + e.Message);
            }
        }
        public void Disconnect()
        {
            sftp.Disconnect();
        }
    }
}
