using MySql.Data.MySqlClient;
using System;
using System.Data.SqlTypes;

namespace Agent2._0
{
    class tblDevice
    {
        public int id { get; set; }
        public string mac { get; set; }
        public string mac2 { get; set; }
        public string sn { get; set; }
        public tblDevice(int id, string mac, string mac2, string sn)
        {
            this.id = id;
            this.mac = mac;
            this.mac2 = mac2;
            this.sn = sn;
        }

        public tblDevice(ServerDatabase db, Excel excel)
        {
            int id = 1;
            string mac = "";
            string mac2 = "";
            string sn = excel.FileSerialNum;
            int countDV = db.Count("SELECT COUNT(id) from tbl_device WHERE sn='" + sn + "'");
            string queryStr;

            if (countDV == 0)
            {
                id = db.GetInt16("SELECT MAX(id) FROM tbl_device") + 1;
            }
            else 
            {
                queryStr = string.Format("SELECT id FROM tbl_device WHERE sn = '{0}' LIMIT 1", sn);
                id = db.GetInt16(queryStr);
            }

            queryStr = string.Format("SELECT mac FROM tbl_import_mac_sn WHERE sn = '{0}'", sn);
            mac = db.GetString(queryStr);
            queryStr = string.Format("SELECT mac2 FROM tbl_import_mac_sn WHERE sn = '{0}'", sn);
            mac2 = db.GetString(queryStr);

            this.id = id;
            this.sn = sn;
            this.mac = mac;
            this.mac2 = mac2;
        }
    }
}
