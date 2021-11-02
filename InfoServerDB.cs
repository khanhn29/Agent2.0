using System;
using System.Xml;

namespace Agent2._0
{
    class InfoServerDB
    {
        public string Ip{get;set;}
        public string Name{get;set;}
        public string Username{get;set;}
        public string Pass{get;set;}
        public string Port{get;set;}
        public InfoServerDB()
        {
            string strExeFilePath = System.Reflection.Assembly.GetEntryAssembly().Location;
            string strWorkPath = System.IO.Path.GetDirectoryName(strExeFilePath);
            XmlTextReader xtr = new XmlTextReader(strWorkPath + "/AgentCfg.xml");
            Console.WriteLine("[Info] Getting MySQL Database info:");
            while (xtr.Read())
            {
                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DBIpAddress")
                {
                    Ip = xtr.ReadElementContentAsString();
                    Console.WriteLine("         IpAddress: " + Ip);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DBName")
                {
                    Name = xtr.ReadElementContentAsString();
                    Console.WriteLine("         Name: " + Name);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DBUsername")
                {
                    Username = xtr.ReadElementContentAsString();
                    Console.WriteLine("         Username: " + Username);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DBPassword")
                {
                    Pass = xtr.ReadElementContentAsString();
                    Console.WriteLine("         Password: " + Pass);
                }

                if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DBPort")
                {
                    Port = xtr.ReadElementContentAsString();
                    Console.WriteLine("         Port: " + Port);
                }
            }
            xtr.Close();
        }
    }
}
