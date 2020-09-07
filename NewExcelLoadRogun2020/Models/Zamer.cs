using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace NewExcelLoadRogun2020.Models
{
    class Zamer
    {
        public ZamerDate zamerDate { get; private set; }
        public PNum projectNum { get; private set; }
        public SensorType sensorType { get; private set; }
        public SensorMeasurements value { get; private set; }

        public Zamer(ZamerDate date, PNum pNum, SensorType sensor, SensorMeasurements measurements)
        {
            zamerDate = date;
            projectNum = pNum;
            sensorType = sensor;
            value = measurements;
        }
        public Zamer(string date, string pNum, string sensor, string measurements)
        {
            zamerDate = ZamerDate.Create(date);
            projectNum = PNum.Create(pNum);
            sensorType = SensorType.Create(sensor);
            value = SensorMeasurements.Create(measurements);

        }
        public static Zamer Create(string date, string pNum, string sensor, string measurements)
        {
            return new Zamer(date, pNum, sensor, measurements);
        }
        public static Zamer Create(string value,string data, Datchik datchik)
        {
            return new Zamer(data, datchik.pNum, datchik.type, value);
        }
        public static explicit operator string (Zamer zamer)
        {
            string retStr = "";
            retStr += zamer.projectNum;
            retStr += " " + zamer.sensorType;
            retStr += " " + zamer.zamerDate;
            retStr += " " + zamer.value + "\n";
            return retStr;
        }

        public static implicit operator ExportToDBClass(Zamer oldValue)
        {
            ExportToDBClass newValue = new ExportToDBClass(oldValue.zamerDate, oldValue.projectNum,
                oldValue.sensorType, oldValue.value);
            return newValue;
        }

        public static List<ExportToDBClass> ToDBList(List<Zamer> oldList)
        {
            List<ExportToDBClass> newList = new List<ExportToDBClass>();

            foreach (var x in oldList)
                newList.Add(x);

            return newList;
        }
    }

    class SensorMeasurements
    {
        private readonly float _value;
        private SensorMeasurements(string str)
        {
            _value = float.Parse(str);
        }

        public static SensorMeasurements Create(string str)
        {
            Regex regex = new Regex(@"(\d+)\.*(\d*)");
            str = str.Replace(",", ".");

            if (string.IsNullOrWhiteSpace(str) || !regex.IsMatch(str))
                throw new Exception("Пустой замер");

            str = regex.Match(str).Value;
            return new SensorMeasurements(str);
        }

        public static implicit operator float(SensorMeasurements measurements)
        {
            return measurements._value;
        }
        public static implicit operator string(SensorMeasurements measurements)
        {
            return measurements._value.ToString();
        }
        
    }

    class SensorType
    {
        private readonly string _value;

        private SensorType(string str)
        {
            _value = str;
        }

        public static SensorType Create(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                throw new Exception("Пустой тип датчика");
            return new SensorType(str);
        }

        public static implicit operator string(SensorType sens)
        {
            return sens._value;
        }
    }

    class PNum
    {
        private readonly string _value;

        private PNum(string str)
        {
            _value = str;
        }

        public static PNum Create(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                throw new Exception("Пустой проектный номер");
            return new PNum(str);
        }

        public static implicit operator string (PNum pNum)
        {
            return pNum._value;
        }
    }

    class ZamerDate
    {
        private readonly DateTime _value;

        private ZamerDate(string DateStr)
        {
            _value = Convert.ToDateTime(DateStr);
        }

        public static ZamerDate Create(string DateStr)
        {
            if (DateStr == null)
                throw new Exception("Пустая дата");

            Regex regex = new Regex(@"(\d{2})\.(\d{2})\.(\d{4})");
            if (regex.IsMatch(DateStr))
            {
                string nowDate = regex.Match(DateStr).Value;
                int year = int.Parse(nowDate.Substring(nowDate.LastIndexOf('.') + 1));
                nowDate = nowDate.Substring(0, nowDate.LastIndexOf('.'));
                int month = int.Parse(nowDate.Substring(nowDate.LastIndexOf('.') + 1));
                int day = int.Parse(nowDate.Substring(0, nowDate.LastIndexOf('.')));
                if (year > 2000 && year < 2030 && month <= 12 && day <= DateTime.DaysInMonth(year, month))
                {
                    return new ZamerDate(DateStr);
                }
                else
                {
                    throw new Exception("Неверный формат даты");
                }
            }
            else if (DateStr.Contains("Нулевой отсчет после перемонтажа"))
            {
                DateStr = "11.11.1112";
            }
            else if (DateStr.Contains("Нулевой отсчет"))
            {
                DateStr = "11.11.1111";
            }

            return new ZamerDate(DateStr);
        }

        public static implicit operator DateTime(ZamerDate date)
        {
            return date._value;
        }
        public static implicit operator string(ZamerDate date)
        {
            return date._value.ToString();
        }
    }
}
