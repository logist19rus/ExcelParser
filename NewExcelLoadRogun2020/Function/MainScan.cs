using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NewExcelLoadRogun2020.Models;
using System.Text.Json;

namespace NewExcelLoadRogun2020.Function
{
    class MainScan
    {
        public static List<Zamer> StartScan(string path)
        {
            var allValues = new List<Zamer>();
            List<string> innerFiles = Directory.GetFiles(path).ToList();

            int s = 0;

            foreach (var x in innerFiles)
            {
                s += ScanFile(x).Count;
                allValues.AddRange(ScanFile(x));
            }

            return allValues;
        }

        public static List<Zamer> MultiThreadScan(string path)
        {
            var allValues = new List<Zamer>();
            List<string> innerFiles = Directory.GetFiles(path).ToList();

            var allFilesInfo = new List<System.IO.FileInfo>();
            foreach (var x in innerFiles)
            {
                allFilesInfo.Add(new System.IO.FileInfo(x));
            }
            allFilesInfo = allFilesInfo.OrderByDescending(x => x.Length).ToList();
            innerFiles.Clear();
            foreach (var x in allFilesInfo)
            {
                innerFiles.Add(x.FullName);
            }
            allFilesInfo.Clear();
            GC.Collect();

            allValues = ThreadController.StartScanTh(innerFiles);

            return allValues;
        }

        public static List<Zamer> ScanFile(string fileName)
        {
            List<Zamer> nowValues = new List<Zamer>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                var worksheetPart = workbookPart.WorksheetParts.ToList();
                for(int i = 0; i < worksheetPart.Count; i++)
                {
                    Models.FileInfo fileInfo = fillFileInfoFromColl(workbookPart, worksheetPart[i].Worksheet);

                    for (int j = 0; j < fileInfo.idValues.Count; j++)
                    {
                        int startScan = 0;
                        var thisCol = getMyCol(workbookPart, worksheetPart[i].Worksheet, int.Parse(fileInfo.idValues[j].value));

                        if (j > 0)
                            fileInfo.startNextRound();

                        for (int k = 0; k < thisCol.Count && startScan < 1; k++)
                        {
                            if (GetRowIdFromXlsCol(thisCol[k].location) == fileInfo.idStartValues)
                                startScan = k;
                            else if (GetRowIdFromXlsCol(thisCol[k].location) == fileInfo.idType)
                                fileInfo.nowDat.type = thisCol[k].value;
                            else if (GetRowIdFromXlsCol(thisCol[k].location) == fileInfo.idPNum)
                                fileInfo.nowDat.pNum = thisCol[k].value;
                            else if (GetRowIdFromXlsCol(thisCol[k].location) == fileInfo.idZeroDate)
                                fileInfo.nowDat.dataZero = getRealDate(thisCol[k].value);
                            else if (GetRowIdFromXlsCol(thisCol[k].location) == fileInfo.idDateReinstall)
                                fileInfo.nowDat.deteReinstall = getRealDate(thisCol[k].value);
                        }
                        if (string.IsNullOrWhiteSpace(fileInfo.nowDat.pNum))
                            continue;
                        fileInfo.fillEmptyValues();
                        //try { fileInfo.fillEmptyValues(); }
                        //catch { continue; }

                        for(int k = startScan+1; k < thisCol.Count; k++)
                        {
                            string thisDate = string.Empty;
                            foreach(var x in fileInfo.dates)
                            {
                                if(GetRowIdFromXlsCol(thisCol[k].location) == GetRowIdFromXlsCol(x.location))
                                {
                                    thisDate = x.value;
                                    if (thisDate == "11.11.1111")
                                        thisDate = fileInfo.nowDat.dataZero;
                                    else if (thisDate == "11.11.1112")
                                        thisDate = fileInfo.nowDat.deteReinstall;
                                    break;
                                }
                            }
                            string thisValue = thisCol[k].value;
                            //try
                            //{
                            //    nowValues.Add(Zamer.Create(thisValue, thisDate, fileInfo.nowDat));
                            //}
                            //catch(Exception ex)
                            //{
                            //    //Console.WriteLine(ex.Message);
                            //}
                            Zamer test = CreateNewZamer(thisValue, thisDate, fileInfo.nowDat);
                            if (test != null)
                                nowValues.Add(test);
                        }
                    }
                }
            }

            return nowValues;
        }

        public static Zamer CreateNewZamer(string thisValue, string thisDate, Datchik nowDat)
        {
            Zamer fRet = null;

            if (nowDat != null)
            {
                if (!string.IsNullOrEmpty(nowDat.pNum))//proectniy nomer
                {
                    Regex regex = new Regex(@"(\d+)\.*(\d*)");
                    thisValue = thisValue.Replace(",", ".");

                    if (!string.IsNullOrWhiteSpace(thisValue) && regex.IsMatch(thisValue))//pokazanye zamera
                    {
                        if (!string.IsNullOrEmpty(nowDat.type))//type datchika
                        {
                            Regex regex2 = new Regex(@"(\d{2})\.(\d{2})\.(\d{4})");
                            if (!string.IsNullOrEmpty(thisDate) && regex2.IsMatch(thisDate))
                            {
                                string nowDate = regex2.Match(thisDate).Value;
                                int year = int.Parse(nowDate.Substring(nowDate.LastIndexOf('.') + 1));
                                nowDate = nowDate.Substring(0, nowDate.LastIndexOf('.'));
                                int month = int.Parse(nowDate.Substring(nowDate.LastIndexOf('.') + 1));
                                int day = int.Parse(nowDate.Substring(0, nowDate.LastIndexOf('.')));
                                if (year > 2000 && year < 2030 && month <= 12 && day <= DateTime.DaysInMonth(year, month))
                                {
                                    fRet = new Zamer(date: thisDate, pNum: nowDat.pNum, sensor: nowDat.type, measurements: thisValue);
                                }
                            }
                        }
                    }
                }
            }

            return fRet;
        }

        public static int GetColIdFromXlsCol(string xlsCol)
        {
            int myId = 0;
            xlsCol = Regex.Match(xlsCol, @"[A-Z]+").Value;

            for(int i = 0; i < xlsCol.Length; i++)
            {
                int timedP = 1;
                for (int j = 0; j < xlsCol.Length - i - 1; j++)
                {
                    timedP *= 26;
                }
                myId += timedP * (xlsCol[i] - 'A' + 1);
            }
            return myId;
        }
        public static int GetRowIdFromXlsCol(string xlsCol)
        {
            int myId = 0;
            xlsCol = Regex.Match(xlsCol, @"\d+").Value;
            myId = int.Parse(xlsCol);
            return myId;
        }

        public static Models.FileInfo fillFileInfoFromColl(WorkbookPart workbookPart, Worksheet sheet)
        {
            var myCol = getMyCol(workbookPart, sheet, GetColIdFromXlsCol("A"));
            Models.FileInfo fileInfo = new Models.FileInfo();
            int startScan = 0;

            for (int i = 0; i < myCol.Count & !fileInfo.Filled(); i++)
            {
                if (fileInfo.idType == 0  && (((string)myCol[i]).ToLower().Contains("анкер") || ((string)myCol[i]).ToLower().Contains("прибора")))
                {
                    fileInfo.idType = GetRowIdFromXlsCol(myCol[i].location);
                }
                else if (((string)myCol[i]).ToLower().Contains("проект"))
                {   
                    fileInfo.idPNum = GetRowIdFromXlsCol(myCol[i].location);
                }
                else if (((string)myCol[i]).ToLower().Contains("замер") || ((string)myCol[i]).ToLower().Contains("(date)"))
                {
                    fileInfo.idStartValues = GetRowIdFromXlsCol(myCol[i].location);
                    startScan = i;
                }
                else if (((string)myCol[i]).ToLower().Contains("установки") && ((string)myCol[i]).ToLower().Contains("да"))
                {
                    fileInfo.idZeroDate = GetRowIdFromXlsCol(myCol[i].location);
                }
                else if (((string)myCol[i]).ToLower().Contains("перемонтажа"))
                {
                    fileInfo.idDateReinstall = GetRowIdFromXlsCol(myCol[i].location);
                }
            }//заполнили все индексы важных строк

            for (int i = startScan + 1; i < myCol.Count; i++)
            {
                if (((string)myCol[i]) != null)
                {
                    myCol[i].value = getRealDate(myCol[i].value);
                    fileInfo.addDates(myCol[i]);
                }
            }//запомнили все существующие на листе даты

            var valuesRowStr = getMyRow(workbookPart, sheet, fileInfo.idStartValues);
            foreach(var x in valuesRowStr)
            {
                if (x.value.ToLower().Contains("показ") || x.value.ToLower().Contains("градус"))
                    fileInfo.idValues.Add(new XlsSqrt(GetColIdFromXlsCol(x.location).ToString(), x.location));
            }//запомнили все индексы колонок с показаниями

            return fileInfo;
        }

        public static List<XlsSqrt> getMyRow(WorkbookPart workbookPart, Worksheet sheet, int idRow)
        {
            List<XlsSqrt> myRow = new List<XlsSqrt>();

            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            var valuesRow = sheet.Descendants<Row>().FirstOrDefault(x => x.RowIndex == idRow);

            foreach (Cell cell in valuesRow)
            {
                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                {
                    int ssid = int.Parse(cell.CellValue.Text);
                    string str = sst.ChildElements[ssid].InnerText;
                    myRow.Add(new XlsSqrt(str, cell.CellReference.InnerText));
                }
                else if (cell.CellValue != null)
                {
                    myRow.Add(new XlsSqrt(cell.CellValue.Text, cell.CellReference.InnerText));
                }
            }

            return myRow;
        }

        public static List<XlsSqrt> getMyCol (WorkbookPart workbookPart, Worksheet sheet, int idRow)
        {
            List<XlsSqrt> myCol = new List<XlsSqrt>();
            string idRowStr = string.Empty;

            do
            {
                char nowCh = (char)('A' + (idRow % 26) - 1);
                idRowStr = nowCh + idRowStr;
                idRow = idRow / 26;
            } while (idRow > 0);

            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;
            string regStr = idRowStr + @"(\d+)";
            var cells = sheet.Descendants<Cell>().Where(x => Regex.Match(x.CellReference.InnerText, @"[A-Z]+").Value == idRowStr).ToList();

            foreach (Cell cell in cells)
            {
                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                {
                    int ssid = int.Parse(cell.CellValue.Text);
                    string str = sst.ChildElements[ssid].InnerText;
                    myCol.Add(new XlsSqrt(str, cell.CellReference.InnerText));
                }
                else if (cell.CellValue != null)
                {
                    myCol.Add(new XlsSqrt(cell.CellValue.Text, cell.CellReference.InnerText));
                }
            }

            return myCol;
        }

        public static List<List<string>> getStrings(WorkbookPart workbookPart, Worksheet sheet)
        {
            var rowsStr = new List<List<string>>();

            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            var rows = sheet.Descendants<Row>();

            Console.WriteLine("Row count = {0}", rows.LongCount());
            foreach (Row row in rows)
            {
                rowsStr.Add(new List<string>());
                foreach (Cell c in row.Elements<Cell>())
                {
                    string str = string.Empty;
                    if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                    {
                        int ssid = int.Parse(c.CellValue.Text);
                        str = sst.ChildElements[ssid].InnerText;
                        if (str.Contains("Нулевой отсчет после перемонтажа"))
                        {
                            str = "11.11.1112";
                        }
                        else if (str.Contains("Нулевой отсчет"))
                        {
                            str = "11.11.1111";
                        }
                    }
                    else if (c.CellValue != null && c.CellFormula != null)
                    {
                        if (c.CellValue.Text.Contains("Нулевой отсчет после перемонтажа"))
                        {
                            str = "11.11.1112";
                        }
                        else if (c.CellValue.Text.Contains("Нулевой отсчет"))
                        {
                            str = "11.11.1111";
                        }
                        else
                        {
                            str = c.CellValue.Text;
                        }
                    }
                    else if (c.CellValue != null && Regex.IsMatch(c.CellReference.InnerText, @"A(\d+)"))
                    {
                        if(Regex.IsMatch(c.CellValue.Text, @"(\d{5})"))
                        {
                            DateTime warWithXML = Convert.ToDateTime("30.12.1899");
                            warWithXML = warWithXML.AddDays(int.Parse(c.CellValue.Text));
                            str = warWithXML.ToString();
                        }
                        else if(Regex.IsMatch(c.CellValue.Text, @"(\d{2})\.(\d{2})\.(\d{4})"))
                        {
                            str = c.CellValue.Text;
                        }
                    }
                    else if (c.CellValue != null)
                    {
                        if (c.CellValue.Text.Contains("Нулевой отсчет после перемонтажа"))
                        {
                            str = "11.11.1112";
                        }
                        else if (c.CellValue.Text.Contains("Нулевой отсчет"))
                        {
                            str = "11.11.1111";
                        }
                        else
                            str = c.CellValue.Text;
                    }
                    else
                    {
                        str = " ";
                    }
                    rowsStr[rowsStr.Count - 1].Add(str);
                }
            }

            return rowsStr;
        }

        public static Models.FileInfo fillFileInfo(List<List<string>> rows, string fileName)
        {
            Models.FileInfo fileInfo = new Models.FileInfo();
            fileInfo.filename = fileName.Substring(fileName.LastIndexOf("\\") + 1);
            for (int i = 0; i < rows.Count & !fileInfo.Filled(); i++)
            {
                if (rows[i][0].ToLower().Contains("анкер") || rows[i][0].ToLower().Contains("прибора"))
                {
                    fileInfo.idType = i;
                }
                else if (rows[i][0].ToLower().Contains("проект"))
                {
                    fileInfo.idPNum = i;
                }
                else if (rows[i][0].ToLower().Contains("замер"))
                {
                    fileInfo.idStartValues = i;
                }
                else if (rows[i][0].ToLower().Contains("установки") & rows[i][0].ToLower().Contains("да"))
                {
                    fileInfo.idZeroDate = i;
                }
                else if (rows[i][0].ToLower().Contains("перемонтажа"))
                {
                    fileInfo.idDateReinstall = i;
                }
            }

            for (int i = fileInfo.idStartValues + 1; i < rows.Count; i++)
                if (rows[i].Count > 0 && Regex.IsMatch(rows[i][0], @"(\d{2})\.(\d{2})\.(\d{4})"))
                    fileInfo.addDates(rows[i][0], "00");

            for (int i = 0; i < rows[fileInfo.idStartValues].Count; i++)
                if (rows[fileInfo.idStartValues][i].Contains("показан"))
                    fileInfo.idValues.Add(new XlsSqrt(i.ToString(), "00"));

            return fileInfo;
        }

        public static string getRealDate(string exlDate)
        {
            DateTime newDate = Convert.ToDateTime("30.12.1899");

            if (exlDate.Length == 5 && Regex.IsMatch(exlDate, @"(\d{5})"))
            {
                newDate = newDate.AddDays(int.Parse(exlDate));
                exlDate = newDate.ToString();
            }

            return exlDate;
        }
    }
}
