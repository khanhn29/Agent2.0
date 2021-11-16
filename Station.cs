using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agent2._0
{
    class Station
    {
        
        string Name { get; set; }
        string FolderName { get; set; }
        Int16 Code { get; set; }
        Type Type { get; set; }
    }
    enum Type
    {
        Device,
        DeviceResult,
        DetailResult
    }
}
