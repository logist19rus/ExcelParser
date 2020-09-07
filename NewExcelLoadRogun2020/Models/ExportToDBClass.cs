using System;
using System.Collections.Generic;
using System.Text;

namespace NewExcelLoadRogun2020.Models
{
    class ExportToDBClass
    {
        public DateTime DateOfMeasurement { get; set; }
        public string ProjectNumber { get; set; }
        public string SensorType { get; set; }
        public float SensorMeasurements { get; set; }

        public ExportToDBClass(DateTime _date, string _pnum, string _type, float _znach)
        {
            DateOfMeasurement = _date;
            ProjectNumber = pnumrepair(_pnum);
            SensorType = replasim(_type);
            SensorMeasurements = _znach;
        }

        public static string replasim(string stroka)
        {
            stroka = stroka.Replace(",", "");
            stroka = stroka.Replace(".", "");
            stroka = stroka.Replace(" ", "");
            stroka = stroka.Replace("x", "X");
            stroka = stroka.Replace("х", "X");
            stroka = stroka.Replace("X", "Х");
            stroka = stroka.Replace("M", "М");
            stroka = stroka.Replace("p", "П");
            stroka = stroka.Replace("P", "П");
            stroka = stroka.Replace("a", "а");
            stroka = stroka.Replace("A", "А");

            return stroka;
        }
        public static string pnumrepair(string stroka)
        {
            stroka = stroka.Replace("-", "");
            stroka = stroka.ToUpper();

            return stroka;
        }

    }
}
