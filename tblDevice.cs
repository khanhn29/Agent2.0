using MySql.Data.MySqlClient;
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
            MySqlDataReader rdr = db.Reader("SELECT MAX(id) FROM tbl_device");
            //Read number of device
            rdr.Read();
            try{
                this.id = rdr.GetInt16(0) + 1;
            }
            catch (SqlNullValueException){
                this.id = 1;
            }
            rdr.Close();

            this.sn = excel.ReadCell(6, 2);
            this.mac = "";
            this.mac2 = "";
        }
    }
}
