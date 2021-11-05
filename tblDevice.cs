﻿using MySql.Data.MySqlClient;
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
        //public tblDevice()
        //{
        //    this.id = 0;
        //    this.mac = "";
        //    this.sn = "";
        //}
        public tblDevice(ServerDatabase db, Excel excel)
        {
            int id = 1;
            string mac = "";
            string mac2 = "";
            string sn = "";
            MySqlDataReader rdr = db.Reader("SELECT MAX(id) FROM tbl_device");

            rdr.Read();
            try{
                id = rdr.GetInt16(0) + 1;
            }
            catch (SqlNullValueException){
                id = 1;
            }
            rdr.Close();

            sn = excel.ReadCell(6, 2);

            rdr = db.Reader("SELECT mac, mac2 FROM tbl_import_mac_sn WHERE sn = '" + sn + "'");
            rdr.Read();
            try{
                mac = rdr.GetString(0);
                mac2 = rdr.GetString(1);
            }
            catch (MySqlException e){
                Log.Error("" + e.Message);
            }
            rdr.Close();

            this.id = id;
            this.sn = sn;
            this.mac = mac;
            this.mac2 = mac2;
        }
    }
}
