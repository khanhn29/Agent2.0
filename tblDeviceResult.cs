using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Data.SqlTypes;
using System.Globalization;

namespace Agent2._0
{
    class tblDeviceResult
    {
        public UInt64 id { get; set; }
        public UInt64 device_id { get; set; }
        public int campaign_id { get; set; }
        public int line { get; set; }
        public string date { get; set; }
        public string time { get; set; }
        public string station_name { get; set; }
        public string tester_name { get; set; }
        public string tester_id { get; set; }
        public int latest { get; set; }
        public string result { get; set; }
        public tblDeviceResult()
        {
            id = 0;
            device_id = 0;
            campaign_id = 0;
            line = 0;
            date = "";
            time = "";
            station_name = "";
            tester_name = "";
            tester_id = "";
            latest = 0;
            result = "";
        }
        public tblDeviceResult(ServerDatabase db, Excel excel, tblCampaign campaign)
        {
            Hashtable stationnameNums = new Hashtable()
            {
                {"TRX-TEST", 1},
                {"PA-TEST", 2},
                {"FILTER-TEST", 3},
                {"ANT-TEST", 4},
                {"ASSEM-RRU", 5},
                {"SAVE-DATA", 6},
                {"PERFORMANCE-TEST", 7},
                {"AIR-TEST", 8},
                {"RRU-BURN-IN", 9},
                {"TEST-TRX-BURN-IN", 10},
                {"AIR-TEST-AFTER-BURN-IN", 11},
                {"THERMAL-CYCLE", 12},
                {"VIBRATION-TEST", 13},
                {"PERFORMANCE-TEST-AFTER-VIBRATION", 14},
                {"PACKAGE", 15}
            };

            MySqlDataReader rdr = db.Reader("SELECT MAX(id) FROM tbl_device_result");
            rdr.Read();
            try{
                id = rdr.GetUInt64(0) + 1;
            }
            catch (SqlNullValueException){
                id = 1;
            }
            rdr.Close();

            string sn = excel.FileSerialNum;
            string queryStr = string.Format("SELECT id FROM tbl_device WHERE sn='{0}' LIMIT 1", sn);
            device_id = db.GetUInt64(queryStr);

            campaign_id = campaign.Id;
            try
            {
                line = Int16.Parse(excel.ReadCell(3, 2));
            }
            catch {
                line = 1;
            }
            date = GetDate(excel.ReadCell(8, 2));
            time = excel.ReadCell(9, 2);
            station_name = stationnameNums[excel.StationName].ToString();
            tester_name = excel.ReadCell(4, 2);
            tester_id = excel.ReadCell(5, 2);
            latest = 1;
            result = excel.FileResult;
        }
        private string GetDate(string rawDate)
        {
            DateTime dt;
            try
            {
                dt = DateTime.ParseExact(rawDate, "d/M/yyyy:H:m", CultureInfo.InvariantCulture);
                return dt.ToString("yyyy-MM-dd H:m:s");
            }
            catch(FormatException e)
            {
                Log.Error("Read date error: " + e.Message);
                return DateTime.Now.ToString("yyyy-MM-dd H:m:s");
            }
        }
    }
}
