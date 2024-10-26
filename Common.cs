using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FranchisorEXE
{
    public class Common
    {

        public static void WriteLog(string msg)
        {
            try
            {
                string filePath = "ProcessLog.txt";

                if (!File.Exists(filePath))
                {
                    File.Create(filePath).Close();

                }
                StreamWriter sw = new StreamWriter(filePath, true);

                //sw.WriteLine(DateTime.Now.ToString());

                sw.WriteLine(msg);

                sw.Flush();
                sw.Close();
            }
            catch { }

        }

    }
}
