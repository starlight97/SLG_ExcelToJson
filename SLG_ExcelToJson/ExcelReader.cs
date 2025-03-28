﻿using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System;

namespace SLG_ExcelToJson
{
    public static class ExcelReader
    {
        public static Application ExcelApp;
        public static Workbooks ExcelBooks;
        public static Workbook ExcelBook;
        public static Sheets ExcelSheets;
        public static Worksheet ExcelSheet;

        public static List<ExcelSheetInfo> InfoList { get; set; }

        public static void Init()
        {
            InfoList = new List<ExcelSheetInfo>();

            ExcelApp = new Application();
            ExcelBooks = ExcelApp.Workbooks;
        }

        public static void AddExcelFile(string filePath)
        {
            ExcelBook = ExcelApp.Workbooks.Add(filePath);
            ExcelSheets = ExcelBook.Sheets;

            //파일 입력 받을때마다 Sheet 개별을 가져옴
            for (int i = 1; i <= ExcelSheets.Count; i++)
            {
                ExcelSheetInfo info = new ExcelSheetInfo();
                InfoList.Add(info);

                try
                {
                    info.ExcelSheet = ExcelSheets.Item[i];
                }
                catch (Exception e)
                {
                    var msg = "Json Convert Error\r\n";
                    foreach (var dataName in info.DataNames)
                    {
                        msg += dataName + "\r\n";
                    }
                    
                    ErrorManager.instance.AddErrorLog(msg);
                    ErrorManager.instance.Show();
                    throw;
                }
            }
        }

        //public static void AddExcelFiles(string[] filePaths)
        //{
        //    foreach (string path in filePaths)
        //        AddExcelFile(path);
        //}

        public static List<List<List<dynamic>>> GetAllSheetValues()
        {
            var rtnList = new List<List<List<dynamic>>>();
            for (int i = 0; i < InfoList.Count; i++)
                rtnList.Add(InfoList[i].GetSheetValues());

            return rtnList;
        }

        public static List<List<dynamic>> GetSheetValuesByIndex(int index)
        {
            return InfoList[index].GetSheetValues();
        }

        public static void Free()
        {
            //저장할지 물어보는거 취소.
            ExcelApp.DisplayAlerts = false;
            ExcelApp.Quit();

            foreach (var info in InfoList)
            {
                info.Free();
            }
            InfoList.Clear();
            Marshal.ReleaseComObject(ExcelSheets);
            Marshal.ReleaseComObject(ExcelBook);
            Marshal.ReleaseComObject(ExcelBooks);
            Marshal.ReleaseComObject(ExcelApp);
        }
    }
}
