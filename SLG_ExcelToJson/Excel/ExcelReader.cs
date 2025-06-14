using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System;

namespace SLG_ExcelToJson
{
    public class ExcelReader
    {
        public static Application ExcelApp;
        public static Workbooks ExcelBooks;
        public static Workbook ExcelBook;
        public static Sheets ExcelSheets;
        public static Worksheet ExcelSheet;

        public static List<ExcelSheetInfo> InfoList => _infoList;
        private static List<ExcelSheetInfo> _infoList = new List<ExcelSheetInfo>(); 

        public static void Init()
        {
            ExcelApp = new Application();
            ExcelBooks = ExcelApp.Workbooks;
        }
        
        public static void Clear()
        {
            //저장할지 물어보는거 취소.
            ExcelApp.DisplayAlerts = false;
            ExcelApp.Quit();

            foreach (var info in _infoList)
            {
                info.Clear();
            }
            _infoList.Clear();
            Marshal.ReleaseComObject(ExcelSheets);
            Marshal.ReleaseComObject(ExcelBook);
            Marshal.ReleaseComObject(ExcelBooks);
            Marshal.ReleaseComObject(ExcelApp);
        }

        public static void AddExcelFile(string filePath)
        {
            ExcelBook = ExcelApp.Workbooks.Add(filePath);
            ExcelSheets = ExcelBook.Sheets;

            //파일 입력 받을때마다 Sheet 개별을 가져옴
            for (int i = 1; i <= ExcelSheets.Count; i++)
            {
                var sheetData = ExcelSheets.Item[i];
                try
                {
                    var name = sheetData.Name;
                    var skipSheet = name.StartsWith("_");
                    if (skipSheet)
                    {
                        continue;
                    }

                    var excelSheet = ExcelSheets.Item[i];
                    var info = new ExcelSheetInfo();
                    info.ExcelSheet = excelSheet;
                    info.RemoveUnUsedData();
                    _infoList.Add(info);
                }
                catch (Exception e)
                {
                    var msg = $"Json Convert Error";
                    foreach (var dataName in sheetData.DataNames)
                    {
                        msg += dataName + "\r\n";
                    }

                    msg += $"{e}\r\n";
                    
                    ErrorManager.instance.AddErrorLog(msg);
                    ErrorManager.instance.Show();
                    throw;
                }
            }
        }
        

        public static List<List<List<dynamic>>> GetAllSheetValues()
        {
            var rtnList = new List<List<List<dynamic>>>();
            for (int i = 0; i < _infoList.Count; i++)
                rtnList.Add(_infoList[i].GetSheetValues());

            return rtnList;
        }

        public static List<List<dynamic>> GetSheetValuesByIndex(int index)
        {
            return _infoList[index].GetSheetValues();
        }
    }
}
