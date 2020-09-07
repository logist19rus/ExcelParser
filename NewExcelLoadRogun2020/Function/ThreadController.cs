using System;
using System.Collections.Generic;
using System.Text.Json;
using NewExcelLoadRogun2020.Models;
using System.Threading;
using System.IO;
using System.Linq;

namespace NewExcelLoadRogun2020.Function
{
    class ThreadController
    {
        const int threadCount = 3;
        public static List<Zamer>[] asd = new List<Zamer>[threadCount];
        
        public static List<Zamer> StartScanTh(List<string> innerFiles)
        {
            var allValues = new List<Zamer>();
            var zaeval = new List<Zamer>();
            Thread[] threads = new Thread[threadCount];
            for(int i = 0; i < threadCount; i++)
            {
                asd[i] = new List<Zamer>();
            }

            for (int i = 0; i < threadCount && i < innerFiles.Count; i++)
            {
                fInfo ddd = new fInfo(innerFiles[i], i);
                threads[i] = new Thread(new ParameterizedThreadStart(GetJsonStr));
                threads[i].Start(ddd);
            }

            for (int i = threadCount; i < innerFiles.Count; i++)
            {
                bool thStarted = false;
                for(int th=0;th< threadCount; th++)
                {
                    if (!threads[th].IsAlive)
                    {
                        fInfo ddd = new fInfo(innerFiles[i], th);
                        threads[th] = new Thread(new ParameterizedThreadStart(GetJsonStr));
                        threads[th].Start(ddd);
                        thStarted = true;
                        break;
                    }
                }
                if (!thStarted)
                {
                    Thread.Sleep(200);
                    i--;
                }
            }
            bool thEndWorking = false;
            while (!thEndWorking)
            {
                thEndWorking = true;
                for (int i = 0; i < threadCount; i++)
                {
                    if (threads[i]!=null && threads[i].IsAlive)
                    {
                        thEndWorking = false;
                        break;
                    }
                }
                Thread.Sleep(100);
            }

            for (int i = 0; i < threadCount; i++)
                allValues.AddRange(asd[i]);

            return allValues;
        }
        public static void GetJsonStr(object filename)
        {
            var nowFile = (fInfo)filename;

            var ass = MainScan.ScanFile(nowFile.fName);

            for(int i = 0; i < ass.Count; i++)
            {
                asd[nowFile.id].Add(ass[i]);
            }
        }
    }
    class fInfo
    {
        public int id { get; set; }
        public string fName { get; set; }

        public fInfo(string _fName, int _id)
        {
            id = _id;
            fName = _fName;
        }
    }
}
