using System;
using System.Collections.Generic;
using System.Text;

namespace NewExcelLoadRogun2020.Models
{
    class FileInfo
    {
        public string filename { get; set; }
        public int idPNum { get; set; } = 0;
        public int idType { get; set; } = 0;
        public int idStartValues { get; set; } = 0;
        public int idZeroDate { get; set; } = 0;
        public int idDateReinstall { get; set; } = 0;
        public Datchik oldDat { get; set; } = null;
        public Datchik nowDat { get; set; } = new Datchik();
        public List<XlsSqrt> dates { get; set; } = new List<XlsSqrt>();
        public List<XlsSqrt> idValues { get; set; } = new List<XlsSqrt>();
        public bool Filled()
        {
            if (idPNum > 0 & idType > 0 & idStartValues > 0 & idZeroDate > 0)
                return true;
            else
                return false;
        }

        public void addDates(string date, string loc)
        {
            if (date.Contains("Нулевой отсчет после перемонтажа"))
                date = "11.11.1112";
            else if (date.Contains("Нулевой отсчет"))
                date = "11.11.1111";

            dates.Add(new XlsSqrt(date, loc));
        }
        public void addDates(XlsSqrt xlsSqrt)
        {
            string date = (string)xlsSqrt;
            if (date.Contains("Нулевой отсчет после перемонтажа"))
                date = "11.11.1112";
            else if (date.Contains("Нулевой отсчет"))
                date = "11.11.1111";

            dates.Add(new XlsSqrt(date, xlsSqrt.location));
        }
        public void startNextRound()
        {
            oldDat = new Datchik();
            oldDat = nowDat;
            nowDat = new Datchik();
        }

        public void fillEmptyValues()
        {
            if (oldDat==null || string.IsNullOrEmpty(oldDat.pNum))
                return;

            if (nowDat.type != null && (nowDat.type.ToLower().Contains("терм") || nowDat.type.ToLower().Contains("4700")) && !(nowDat.pNum.ToUpper().EndsWith("T") || nowDat.pNum.ToUpper().EndsWith("Т")))
            {
                nowDat.pNum += "T";
            }
            else if(nowDat.type == null && oldDat.type != null)
            {
                nowDat.type = oldDat.type;
            }
            if (string.IsNullOrEmpty(nowDat.dataZero))
            {
                if (oldDat != null)
                    nowDat.dataZero = oldDat.dataZero;
                else
                    nowDat.dataZero = null;
            }
            if (string.IsNullOrEmpty(nowDat.deteReinstall))
            {
                if (oldDat != null)
                    nowDat.deteReinstall = oldDat.deteReinstall;
                else
                    nowDat.deteReinstall = null;
                
            }
        }
    }
    class XlsSqrt
    {
        public string value { get; set; }
        public string location { get; set; }
        public XlsSqrt(string _val, string _loc)
        {
            value = _val;
            location = _loc;
        }
        public static explicit operator string(XlsSqrt xls)
        {
            return xls.value;
        }
    }
    class Datchik
    {
        public string pNum { get; set; }
        public string type { get; set; }
        public string dataZero { get; set; }
        public string deteReinstall { get; set; }
    }
}
