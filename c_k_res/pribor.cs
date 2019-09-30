using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace c_k_res
{
    class Pribor
    {
        public Pribor(string s)
        {
            _path = s;
            _number = s;            
            _number = _number.Remove(0, Environment.CurrentDirectory.Length + 1);
        }        
        public string GetNumber()
        {
            return _number;
        }
        public string GetPath()
        {
            return _path;
        }

        private string _number;
        private string _path;        
    }
}
