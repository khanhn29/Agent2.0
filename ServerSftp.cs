using Renci.SshNet;
using Renci.SshNet.Common;
using Renci.SshNet.Sftp;
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
            return @"\sftp\report";
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
                Log.Error(" Connect sftp error: " + e.Message);
            }
        }
        public void Disconnect()
        {
            sftp.Disconnect();
        }
        public void CreateDirectoryRecursively(string path)
        {
            string current = "";

            if (path[0] == '/')
            {
                path = path.Substring(1);
            }

            while (!string.IsNullOrEmpty(path))
            {
                int p = path.IndexOf('/');
                current += '/';
                if (p >= 0)
                {
                    current += path.Substring(0, p);
                    path = path.Substring(p + 1);
                }
                else
                {
                    current += path;
                    path = "";
                }

                try
                {
                    SftpFileAttributes attrs = this.sftp.GetAttributes(current);
                    if (!attrs.IsDirectory)
                    {
                        throw new Exception("not directory");
                    }
                }
                catch (SftpPathNotFoundException)
                {
                    this.sftp.CreateDirectory(current);
                }
            }
        }
    }
}
