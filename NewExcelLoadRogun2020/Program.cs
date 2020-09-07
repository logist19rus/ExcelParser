using System;
using NewExcelLoadRogun2020.Models;
using NewExcelLoadRogun2020.Function;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace NewExcelLoadRogun2020
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\logist\Data\NewRogun2020\all\";
            string pathTest = @"C:\logist\Data\NewRogun2020\";
            var zamery = MainScan.MultiThreadScan(pathTest);

            List<ExportToDBClass> exports = Zamer.ToDBList(zamery);

            //zamery.Clear();

         
             Console.ReadLine();
        }
    }
}
