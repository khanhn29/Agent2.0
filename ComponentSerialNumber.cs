using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agent2._0
{
    class ComponentsSerialNumber
    {
        public string FileName { get; set; }
        public string sn_rru { get; set; }
        public string sn_trx {get; set;}
        public string sn_pa1 { get; set; }
        public string sn_pa2 { get; set; }
        public string sn_fil { get; set; }
        public string sn_ant { get; set; }
        public string mac { get; set; }
        public string mac2 { get; set; }
        public ComponentsSerialNumber()
        {
            FileName = "";
            sn_rru = "";
            sn_trx = "";
            sn_pa1 = "";
            sn_pa2 = "";
            sn_fil = "";
            sn_ant = "";
            mac = "";
            mac2 = "";
        }
        public void Print()
        {
            string msg = string.Format("rru:{0} trx:{1} pa1:{2} pa2:{3} fil:{4} ant:{5} mac1:{6} mac2:{7}",
                sn_rru, sn_trx, sn_pa1, sn_pa2, sn_fil, sn_ant, mac, mac2);
            Log.Info(msg);
        }
    }
}
