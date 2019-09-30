using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace c_k_res
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> dirs = new List<string>(Directory.EnumerateDirectories(Environment.CurrentDirectory));
            List<Pribor> pribors = new List<Pribor>();            

            foreach (string s in dirs)
            {
                pribors.Add(new Pribor(s));
            }

            if (dirs.Count==0)
            {
                Console.WriteLine("В каталоге отсутствуют папки");
            }
            else
            {
                Data data = new Data();
                data.get_data_c2_29_0125W_kOhm(pribors);
                data.get_data_c2_29_0125W_Ohm(pribors);
                data.get_data_c2_29_025W_Ohm(pribors);
                data.get_data_c5_60_1W_Ohm(pribors);
                data.get_data_c5_60_0125W_kOhm(pribors);
                data.print_data();     
            }  

        }
    }
}
