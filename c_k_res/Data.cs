using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace c_k_res
{
    class Data
    {
        public Data ()
        {
            _errors = 0;
            _pki = new Pki();
            _p_e_numbers = new List<string>();
        }


        void c2_29_0125W_kOhm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            switch (s)
            {
                case "18,70":
                    s = "18,7";
                    break;
                case "75":
                    s = "75,0";
                    break;
                case "6,9":
                    s = "6,90";
                    break;
                case "7,5":
                    s = "7,50";
                    break;
                case "9,2":
                    s = "9,20";
                    break;
                case "10":
                    s = "10,0";
                    break;
                case "20":
                    s = "20,0";
                    break;
                default:
                    break;
            }

            if (_pki.c2_29_0125W_kOhm.ContainsKey(s))
            {
                _pki.c2_29_0125W_kOhm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t"+res+"\t" + s);
            }            
        }
        public void get_data_c2_29_0125W_kOhm(List<Pribor> pribors)
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {
                        
                        switch (s)
                        {
                            case "[R_RegRez1]":
                                string res = s;                                
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez3]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez5]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez6]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez8]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez10]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez11]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez13]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez15]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez19]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez23]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez27]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez30]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez31]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez34]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez35]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez38]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez39]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_kOhm_counter(s, res, pribor);
                                break;

                        }
                    }
                }
            }
        } //18 резисторов 

        void c2_29_0125W_Ohm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            switch (s)
            {        
                case "10":
                    s = "10,0";
                    break;
                case "20":
                    s = "20,0";
                    break;
                case "24":
                    s = "24,0";
                    break;
                case "32":
                    s = "32,0";
                    break;
                default:
                    break;
            }

            if (_pki.c2_29_0125W_Ohm.ContainsKey(s))
            {
                _pki.c2_29_0125W_Ohm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t" + res + "\t" + s);
            }
        }
        public void get_data_c2_29_0125W_Ohm(List<Pribor> pribors) //9 резисторов
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {

                        switch (s)
                        {
                            case "[R_RegRez2]":
                                string res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez4]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez7]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez9]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez12]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez14]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez18]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez22]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez26]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_0125W_Ohm_counter(s, res, pribor);
                                break;



                        }
                    }
                }
            }
        }


        void c2_29_025W_Ohm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            switch (s)
            {
                case "47":
                    s = "47,0";
                    break;
                case "75":
                    s = "75,0";
                    break;
                case "92":
                    s = "92,0";
                    break;
                
                default:
                    break;
            }

            if (_pki.c2_29_025W_Ohm.ContainsKey(s))
            {
                _pki.c2_29_025W_Ohm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t" + res + "\t" + s);
            }
        }
        public void get_data_c2_29_025W_Ohm(List<Pribor> pribors) //3 резисторов
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {

                        switch (s)
                        {
                            case "[R_RegRez16]":
                                string res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_29_025W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez20]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_29_025W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez24]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_29_025W_Ohm_counter(s, res, pribor);
                                break;

                        }
                    }
                }
            }
        }



        void c2_33_1W_Ohm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            if (_pki.c2_33_1W_Ohm.ContainsKey(s))
            {
                _pki.c2_33_1W_Ohm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t" + res + "\t" + s);
            }
        }
        public void get_data_c2_33_1W_Ohm(List<Pribor> pribors) //3 резисторов
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {

                        switch (s)
                        {
                            case "[R_RegRez17]":
                                string res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c2_33_1W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez21]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c2_33_1W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez25]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c2_33_1W_Ohm_counter(s, res, pribor);
                                break;

                        }
                    }
                }
            }
        }


        void c5_60_1W_Ohm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            if (_pki.c5_60_1W_Ohm.ContainsKey(s))
            {
                _pki.c5_60_1W_Ohm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t" + res + "\t" + s);
            }
        }
        public void get_data_c5_60_1W_Ohm(List<Pribor> pribors) //3 резисторов
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {

                        switch (s)
                        {
                            case "[R_RegRez28]":
                                string res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c5_60_1W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez32]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c5_60_1W_Ohm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez36]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c5_60_1W_Ohm_counter(s, res, pribor);
                                break;

                        }
                    }
                }
            }
        }



        void c5_60_0125W_kOhm_counter(string s, string res, Pribor pribor)
        {
            s = s.Replace('.', ',');
            switch (s)
            {
                case "1,6":
                    s = "1,60";
                    break;
                case "1,8":
                    s = "1,80";                                    
                    break;

                default:
                    break;
            }
            if (_pki.c5_60_0125W_kOhm.ContainsKey(s))
            {
                _pki.c5_60_0125W_kOhm[s]++;
            }
            else
            {
                _errors++;
                _p_e_numbers.Add(pribor.GetNumber() + "\t" + res + "\t" + s);
            }
        }
        public void get_data_c5_60_0125W_kOhm(List<Pribor> pribors) //3 резисторов
        {
            if (this._p_count == 0) { _p_count = pribors.Count; }

            foreach (Pribor pribor in pribors)
            {
                string s;

                var VD_reg = new FileInfo(pribor.GetPath() + "\\Регулировка\\VD_Reg.ini");

                using (StreamReader sr = VD_reg.OpenText())
                {
                    while ((s = sr.ReadLine()) != null)
                    {

                        switch (s)
                        {
                            case "[R_RegRez29]":
                                string res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("X="), 2);
                                c5_60_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez33]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Y="), 2);
                                c5_60_0125W_kOhm_counter(s, res, pribor);
                                break;
                            case "[R_RegRez37]":
                                res = s;
                                s = sr.ReadLine();
                                s = s.Remove(s.IndexOf("Z="), 2);
                                c5_60_0125W_kOhm_counter(s, res, pribor);
                                break;

                        }
                    }
                }
            }
        }

        /// Вывод

        public void print_data()
        {
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ",";
            try
            {
                FileStream fs = new FileStream(Environment.CurrentDirectory + "\\Output.xls", FileMode.Create);
                using (StreamWriter sw = new StreamWriter(fs, Encoding.Default))
                {

                    sw.WriteLine("Всего приборов:\t" + _p_count + '\n');
                    sw.WriteLine("ОСС2-29В-0,125 кОм\tКол-во устан.\t Коэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c2_29_0125W_kOhm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("ОСС2-29В-0,125 Ом\tКол-во устан.\tКоэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c2_29_0125W_Ohm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("ОСС2-29В-0,25 Ом\tКол-во устан.\tКоэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c2_29_025W_Ohm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("ОСС2-33Н-1 Ом\tКол-во устан.\tКоэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c2_33_1W_Ohm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("C5-60-1 Ом\tКол-во устан.\tКоэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c5_60_1W_Ohm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("C5-60-0,125 кОм\tКол-во устан.\tКоэф. применяемости");
                    foreach (KeyValuePair<string, int> entry in _pki.c5_60_0125W_kOhm)
                    {
                        sw.Write(entry.Key + '\t' + entry.Value + '\t' + Math.Round(Convert.ToDouble(entry.Value, nfi) / Convert.ToDouble(_p_count, nfi), 1) + '\n');
                    }
                    sw.WriteLine("Ошибок:\t" + _errors);
                    foreach (string s in _p_e_numbers)
                    {
                        sw.WriteLine(s);
                    }
                }

                Console.WriteLine("--------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!--------------");
                Console.WriteLine();
                Console.WriteLine("Файл Output.xls успешно создан.");
                Console.WriteLine();
                Console.WriteLine("--------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!--------------");
                Console.ReadLine();

            }
            catch
            {
                Console.WriteLine("--------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!--------------");
                Console.WriteLine();
                Console.WriteLine("Ошибка записи файла. Возможно файл Output.xls уже используется.");
                Console.WriteLine();
                Console.WriteLine("--------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!--------------");
                Console.ReadLine();                
            }
            
        }


        private int _errors;
        private int _p_count;
        private List<string> _p_e_numbers;
        Pki _pki;
    }
}
