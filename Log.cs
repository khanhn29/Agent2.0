using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Agent2._0
{
    class Log
    {
        public static void Info(string text,
                        [CallerFilePath] string file = "",
                        [CallerMemberName] string member = "",
                        [CallerLineNumber] int line = 0)
        {
            Console.WriteLine("[Info]  {0}_{1}({2}): {3}", Path.GetFileName(file), member, line, text);
        }
        public static void Debug(string text,
                [CallerFilePath] string file = "",
                [CallerMemberName] string member = "",
                [CallerLineNumber] int line = 0)
        {
            if(Program.DebugMode)
                Console.WriteLine("[Debug]  {0}_{1}({2}): {3}", Path.GetFileName(file), member, line, text);
        }
        public static void Error(string text,
                        [CallerFilePath] string file = "",
                        [CallerMemberName] string member = "",
                        [CallerLineNumber] int line = 0)
        {
            Console.WriteLine("[Error] {0}_{1}({2}): {3}", Path.GetFileName(file), member, line, text);
        }
        public static void Warning(string text,
                        [CallerFilePath] string file = "",
                        [CallerMemberName] string member = "",
                        [CallerLineNumber] int line = 0)
        {
            Console.WriteLine("[Warning] {0}_{1}({2}): {3}", Path.GetFileName(file), member, line, text);
        }
    }
}
