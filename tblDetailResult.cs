using MySql.Data.MySqlClient;
using System;
using System.Data.SqlTypes;

namespace Agent2._0
{
    class tblDetailResult
    {
        public int id { get; set; }
        public int device_result_id { get; set; }
        public int sequence { get; set; }
        public string item_name { get; set; }
        public string min_value { get; set; }
        public string reading_value { get; set; }
        public string max_value { get; set; }
        public string test_time { get; set; }
        public string result { get; set; }

        public tblDetailResult()
        {
            id = 0;
            device_result_id = 0;
            sequence = 0;
            item_name = "";
            min_value = "";
            reading_value = "";
            max_value = "";
            test_time = "";
            result = "";
        }
        public tblDetailResult(ServerDatabase db, Excel excel, int excelRow, int newDeviceResultId)
        {
            MySqlDataReader rdr = db.Reader("SELECT MAX(id) FROM tbl_detail_result");
            rdr.Read();
            try{
                id = rdr.GetInt16(0) + 1;
            }
            catch(SqlNullValueException){
                id = 1;
            }
            rdr.Close();
            device_result_id = newDeviceResultId;
            sequence = Int16.Parse(excel.ReadCell(excelRow, 1)); ;
            item_name = excel.ReadCell(excelRow, 2);
            min_value = excel.ReadCell(excelRow, 3);
            reading_value = excel.ReadCell(excelRow, 4);
            max_value = excel.ReadCell(excelRow, 5);
            test_time = excel.ReadCell(excelRow, 6);
            result = excel.ReadCell(excelRow, 7);
        }
    }
}
