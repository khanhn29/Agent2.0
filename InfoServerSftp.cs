using System;
using System.Xml;

namespace Agent2._0
{
    class InfoServerSftp
    {
        public string Ip { get; set; }
        public string Username { get; set; }
        public string Pass { get; set; }

        public InfoServerSftp()
        {
            string strExeFilePath = System.Reflection.Assembly.GetEntryAssembly().Location;
            string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);
            XmlTextReader xtr = new XmlTextReader(strWorkPath + "/AgentCfg.xml");
            while (xtr.Read())
            {
                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "SftpIPAddress")
                {

                    this.Ip = xtr.ReadElementContentAsString();
                    Console.WriteLine("SftpIPAddress : " + Ip);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "SftpUsername")
                {

                    this.Username = xtr.ReadElementContentAsString();
                    Console.WriteLine("SftpUsername : " + Username);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "SftpPassword")
                {

                    this.Pass = xtr.ReadElementContentAsString();
                    Console.WriteLine("SftpPassword : " + Pass);
                }

            }
            xtr.Close();
        }
    }
}
