﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;

namespace SLG_ExcelToJson
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        public List<FileManager> FileManagerList;


        private List<string> excelPaths;
        private string currentDirectory;
        private string currentFileName;


        public MainForm()
        {
            FileManagerList = new List<FileManager>();

            excelPaths = new List<string>();

            InitializeComponent();


        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void mbtClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mbtDirectoryOpen_Click(object sender, EventArgs e)
        {
            //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //dialog.IsFolderPicker = true; // true : 폴더 선택 / false : 파일 선택

            //if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            //{
            //    lbxExcelList.Items.Clear();
            //    excelPaths.Clear();
            //    txtSysMsg.Text = dialog.FileName;
            //    currentDirectory = dialog.FileName;
            //    System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(dialog.FileName);
            //    foreach (var file in di.GetFiles())
            //    {
            //        if (file.Extension == ".xlsx")
            //        {
            //            lbxExcelList.Items.Add(file.Name);
            //            excelPaths.Add(file.FullName);
            //        }

            //    }
            //}
        }

        private void mbtConvert_Click(object sender, EventArgs e)
        {
            ResultTextBox.Text = "변환 시작!!! 로딩중.....";

            //if (lbxExcelList.SelectedItems.Count < 1)
            //{
            //    MessageBox.Show("변환할 파일이 없습니다.", "아이고...", MessageBoxButtons.OK,
            //        MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            //    ResultTextBox.Text = "변환 준비중...";
            //    return;
            //}
            //

            ExcelReader.Init();
            ExcelReader.AddExcelFile(currentFileName);

            foreach(var info in ExcelReader.InfoList)
            {
                var fileManager = new FileManager();
                fileManager.FileFullPath = currentDirectory + info.ExcelSheet.Name;
                fileManager.NewFileExtension = ".json";
                Console.WriteLine("fileManager.FileFullPath : "  + fileManager.FileFullPath);
                FileManagerList.Add(fileManager);

            }

            //엑셀파일 저장.
            var allSheetsValues = ExcelReader.GetAllSheetValues();
            for (int i = 0; i < allSheetsValues.Count; i++)
            {
                string sheetText = JsonChanger.ChangToJArrayToString(ExcelReader.InfoList[i].DataNames, allSheetsValues[i]);
                FileManagerList[i].SaveNewFile(sheetText);

                //cs파일 생성
                ClassMaker maker = new ClassMaker(FileManagerList[i].NewFilePath, FileManagerList[i].NewFileName);
                maker.AddField(ExcelReader.InfoList[i].DataNames, ExcelReader.InfoList[i].DataTypeCodes);
                maker.GenerateCSharpCode();

            }
            FileManagerList.Clear();
            ExcelReader.Free();

            ResultTextBox.Text = "변환이 완료되었습니다!!!";
            Process.Start(currentDirectory);
        }


        private void Btn_FileSelected_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = false; // true : 폴더 선택 / false : 파일 선택
            dialog.Filters.Add(new CommonFileDialogFilter("Excel 파일", "*.xlsx")); // 필터 추가

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSysMsg.Text = dialog.FileName;
                currentFileName = dialog.FileName;
                currentDirectory = Path.GetDirectoryName(currentFileName) + "\\";
            }
        }
    }
}
