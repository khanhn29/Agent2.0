using MySql.Data.MySqlClient;
using System;
using System.Data.SqlTypes;

namespace Agent2._0
{
    class ServerDatabase
    {
        public MySqlConnection conn { get; set; }
        public ServerDatabase(InfoServerDB db)
        {
            var stringBuilder = new MySqlConnectionStringBuilder();
            stringBuilder["Server"] = db.Ip;
            stringBuilder["Database"] = db.Name;
            stringBuilder["User Id"] = db.Username;
            stringBuilder["Password"] = db.Pass;
            stringBuilder["Port"] = db.Port;
            String sqlConnectionString = stringBuilder.ToString();
            this.conn = new MySqlConnection(sqlConnectionString);
            try
            {
                this.conn.Open();
            }
            catch(Exception e)
            {
                Log.Error("Create server database error: " + e.Message);
            }
            
        }
        public void Close()
        {
            this.conn.Close();
        }
        public MySqlDataReader Reader(string queryString)
        {
            try
            {
                using var command = new MySqlCommand(queryString, this.conn);

                MySqlDataReader rdr = command.ExecuteReader();

                return rdr;
            }
            catch (Exception e)
            {
                Log.Error("Exception" + e.Message);
                return null;
            }
        }
        public bool InsertDevice(tblDevice dv)
        {
            MySqlDataReader rdr = this.Reader("SELECT COUNT(id) FROM tbl_device WHERE sn='" + dv.sn + "'");
            Int32 nDevice = 0;
            rdr.Read();
            try
            {
                nDevice = rdr.GetInt16(0);
            }
            catch (SqlNullValueException)
            {
                nDevice = 0;
            }
            rdr.Close();
            if (nDevice > 0)
            {
                //RRU sn already in tbl_device
                string insertQuery = "DELETE FROM tbl_device WHERE sn=?sn";
                MySqlCommand cmd = new MySqlCommand(insertQuery, conn);
                cmd.Parameters.Add("?sn", MySqlDbType.String).Value = dv.sn;
                cmd.ExecuteNonQuery();
            }

            if (dv.sn.StartsWith("RRU"))
                return InsertDeviceRRU(dv);
            else
                return InsertDeviceComponent(dv);
        }
        private bool InsertDeviceRRU(tblDevice dv)
        {
            bool ret = false;
            try
            {
                string insertQuery = "INSERT INTO tbl_device(id, mac, mac2, sn, rru_sn) VALUE(?id, ?mac, ?mac2, ?sn, ?rru_sn)";
                MySqlCommand cmd = new MySqlCommand(insertQuery, this.conn);
                cmd.Parameters.Add("?id", MySqlDbType.Int16).Value = dv.id;
                cmd.Parameters.Add("?mac", MySqlDbType.VarChar).Value = dv.mac;
                cmd.Parameters.Add("?mac2", MySqlDbType.VarChar).Value = dv.mac2;
                cmd.Parameters.Add("?sn", MySqlDbType.VarChar).Value = dv.sn;
                cmd.Parameters.Add("?rru_sn", MySqlDbType.VarChar).Value = dv.sn;
                if (cmd.ExecuteNonQuery() == 1)
                {
                    Log.Info("Insert Device: " + dv.id + "," + dv.mac + "," + dv.sn);
                    ret = true;
                }
                else
                {
                    Log.Error("Insert Device Data failed: " + dv.id + "," + dv.mac + "," + dv.sn);
                    ret = false;
                }
            }
            catch (MySqlException ex)
            {
                Log.Error("Insert Device: " + ex.Message);
                ret = false;
            }
            return ret;
        }
        private bool InsertDeviceComponent(tblDevice dv)
        {
            try
            {
                string insertQuery = "INSERT INTO tbl_device(id, mac, mac2, sn, rru_sn) VALUE(?id, ?mac, ?mac2, ?sn, ?rru_sn)";
                MySqlCommand cmd = new MySqlCommand(insertQuery, this.conn);
                cmd.Parameters.Add("?id", MySqlDbType.Int16).Value = dv.id;
                cmd.Parameters.Add("?mac", MySqlDbType.VarChar).Value = "";
                cmd.Parameters.Add("?mac2", MySqlDbType.VarChar).Value = "";
                cmd.Parameters.Add("?sn", MySqlDbType.VarChar).Value = dv.sn;
                cmd.Parameters.Add("?rru_sn", MySqlDbType.VarChar).Value = "";
                if (cmd.ExecuteNonQuery() == 1)
                {
                    Log.Info("Insert Device: " + dv.id + "," + dv.mac + "," + dv.sn);
                    return true;
                }
                else
                {
                    Log.Error("Insert Device Data failed: " + dv.id + "," + dv.mac + "," + dv.sn);
                    return false;
                }
            }
            catch (MySqlException ex)
            {
                Log.Error("Insert Device: " + ex.Message);
                return false;
            }
        }
        public bool InsertDeviceResult(tblDeviceResult dvR)
        {
            try
            {
                string insertQuery = "INSERT INTO tbl_device_result(id, device_id, campaign_id, line, date, time, station_name, tester_name, tester_id, latest, result)" +
                    "VALUE(?id, ?device_id, ?campaign_id, ?line, ?date, ?time, ?station_name, ?tester_name, ?tester_id, ?latest, ?result)";
                MySqlCommand cmd = new MySqlCommand(insertQuery, this.conn);
                cmd.Parameters.Add("?id", MySqlDbType.Int16).Value = dvR.id;
                cmd.Parameters.Add("?device_id", MySqlDbType.Int16).Value = dvR.device_id;
                cmd.Parameters.Add("?campaign_id", MySqlDbType.Int16).Value = dvR.campaign_id;
                cmd.Parameters.Add("?line", MySqlDbType.Int16).Value = dvR.line;
                cmd.Parameters.Add("?date", MySqlDbType.VarChar).Value = dvR.date;
                cmd.Parameters.Add("?time", MySqlDbType.VarChar).Value = dvR.time;
                cmd.Parameters.Add("?station_name", MySqlDbType.VarChar).Value = dvR.station_name;
                cmd.Parameters.Add("?tester_name", MySqlDbType.VarChar).Value = dvR.tester_name;
                cmd.Parameters.Add("?tester_id", MySqlDbType.VarChar).Value = dvR.tester_id;
                cmd.Parameters.Add("?latest", MySqlDbType.Int16).Value = dvR.latest;
                cmd.Parameters.Add("?result", MySqlDbType.VarChar).Value = dvR.result;
                if (cmd.ExecuteNonQuery() == 1)
                {
                    Log.Info("Insert DeviceResult: " + dvR.id + "," + dvR.device_id + "," + dvR.campaign_id + "," + dvR.line + "," + dvR.date + "," + dvR.time + "," + dvR.station_name + "," + dvR.tester_name + "," + dvR.tester_id + "," + dvR.latest + "," + dvR.result);
                    return true;
                }
                else
                {
                    Log.Error("Data not inserted: " + dvR.id + "," + dvR.device_id + "," + dvR.campaign_id + "," + dvR.line + "," + dvR.date + "," + dvR.time + "," + dvR.station_name + "," + dvR.tester_name + "," + dvR.tester_id + "," + dvR.latest + "," + dvR.result);
                    return false;
                }
            }
            catch (MySqlException ex)
            {
                Log.Error("Insert Device Result mysql row: " + ex.Message);
                return false;
            }
        }
        public bool InsertDetailResult(tblDetailResult dtR)
        {
            try
            {
                string insertQuery = "INSERT INTO tbl_detail_result(id, device_result_id, sequence, item_name, min_value, reading_value, max_value, test_time, result)" +
                    "VALUE(?id, ?device_result_id, ?sequence, ?item_name, ?min_value, ?reading_value, ?max_value, ?test_time, ?result)";
                MySqlCommand cmd = new MySqlCommand(insertQuery, this.conn);
                cmd.Parameters.Add("?id", MySqlDbType.Int16).Value = dtR.id;
                cmd.Parameters.Add("?device_result_id", MySqlDbType.Int16).Value = dtR.device_result_id;
                cmd.Parameters.Add("?sequence", MySqlDbType.Int16).Value = dtR.sequence;
                cmd.Parameters.Add("?item_name", MySqlDbType.VarChar).Value = dtR.item_name;
                cmd.Parameters.Add("?min_value", MySqlDbType.VarChar).Value = dtR.min_value;
                cmd.Parameters.Add("?reading_value", MySqlDbType.VarChar).Value = dtR.reading_value;
                cmd.Parameters.Add("?max_value", MySqlDbType.VarChar).Value = dtR.max_value;
                cmd.Parameters.Add("?test_time", MySqlDbType.VarChar).Value = dtR.test_time;
                cmd.Parameters.Add("?result", MySqlDbType.VarChar).Value = dtR.result;
                if (cmd.ExecuteNonQuery() == 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (MySqlException ex)
            {
                Log.Error("Insert Detail Result mysql: " + ex.Message);
                return false;
            }
        }

        public Int32 Count(string queryString)
        {
            Int32 ret = 0;

            try
            {
                using var command = new MySqlCommand(queryString, this.conn);

                MySqlDataReader rdr = command.ExecuteReader();

                ret = rdr.GetInt32(0);
            }
            catch (Exception e)
            {
                Log.Error("Exception" + e.Message);
            }

            return ret;
        }

    }
}
