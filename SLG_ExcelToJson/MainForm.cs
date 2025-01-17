﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Reflection;

namespace SLG_ExcelToJson
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        public List<FileManager> FileManagerList;


        private SaveManager _saveManager;
        private List<string> excelPaths;
        private string currentDirectory;
        private string currentFileName;
        private string currentFileFullPath;

        private string saveTargetDirectory = string.Empty;

        public MainForm()
        {
            FileManagerList = new List<FileManager>();
            _saveManager = new SaveManager();

            excelPaths = new List<string>();

            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var fileName = "settings.txt";
            // 파일이 존재하는지 확인
            if (!File.Exists(fileName))
            {
                // 파일이 없는 경우에는 새로운 파일을 생성
                CreateSettingsFile(fileName);
            }
            string[] settingValue = File.ReadAllLines("settings.txt");


            for (int index=0; index < settingValue.Count(); index++)
            {
                if (settingValue[index] == string.Empty)
                    continue;

                if(index == 0)
                {
                    currentFileFullPath = settingValue[0]; // 첫 번째 줄은 currentFileFullPath
                    txtSysMsg.Text = currentFileFullPath;
                    currentDirectory = Path.GetDirectoryName(currentFileFullPath) + "\\";
                    currentFileName = Path.GetFileNameWithoutExtension(currentFileFullPath);
                }

                if(index == 1)
                {
                    saveTargetDirectory = settingValue[1]; // 두 번째 줄은 saveTargetDirectory
                }
            }          
        }

        private void mbtClose_Click(object sender, EventArgs e)
        {
            string targetFilePath = currentFileFullPath;
            string targetSaveDirectoryPath = saveTargetDirectory;

            string settingValue = targetFilePath + "\r\n";
            settingValue += targetSaveDirectoryPath + "\r\n";
            File.WriteAllText($"settings.txt", settingValue);

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

            //if (Directory.Exists(currentDirectory + "\\json") == false)
            //{
            //    Directory.CreateDirectory(currentDirectory + "\\json");
            //}
            //if (Directory.Exists(currentDirectory + "\\cs") == false)
            //{
            //    Directory.CreateDirectory(currentDirectory + "\\cs");
            //}

            ErrorManager.instance.Init();

            //if (lbxExcelList.SelectedItems.Count < 1)
            //{
            //    MessageBox.Show("변환할 파일이 없습니다.", "아이고...", MessageBoxButtons.OK,
            //        MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            //    ResultTextBox.Text = "변환 준비중...";
            //    return;
            //}
            //

            if (currentFileFullPath == null || File.Exists(currentFileFullPath) == false)
            {
                MessageBox.Show("변환할 파일이 없습니다.", "아이고...", MessageBoxButtons.OK,MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                ResultTextBox.Text = "변환 준비중...";
                return;
            }

            ExcelReader.Init();
            ExcelReader.AddExcelFile(currentFileFullPath);
            
            _saveManager.Init(saveTargetDirectory);
            _saveManager.Save(ExcelReader.InfoList, false);
            // var dataDic = new Dictionary<string, JArray>();
            // foreach (var info in ExcelReader.InfoList)
            // {
            //     var jArray = new JArray();
            //     var jObj = new JObject();
            //
            //     // 데이터 타입 ex(int , string)
            //     jObj = ChangeToJObject(info.DataNames, info.DataTypeNames);
            //     jArray.Add(jObj);
            //
            //     // 데이터 값 ex(1, "홍길동")
            //     foreach (var values in info.DataValues)
            //     {
            //         var jobj = ChangeToJObject(info.DataNames, values, info.DataTypeNames);
            //         if (jobj != null)
            //             jArray.Add(jobj);
            //     }
            //     dataDic.Add(info.ExcelSheet.Name, jArray);
            // }
            

            // json 여러개로 뽑을때 사용
            // 엑셀파일 저장.
            //var allSheetsValues = ExcelReader.GetAllSheetValues();
            //for (int i = 0; i < allSheetsValues.Count; i++)
            //{
            //    JsonChanger.ChangToJArrayToString(ExcelReader.InfoList[i].DataNames, allSheetsValues[i]);
            //    string sheetText = JsonChanger.ChangToJArrayToString(ExcelReader.InfoList[i].DataNames, allSheetsValues[i]);

            //    FileManagerList[i].SaveNewFile_Temp(sheetText);

            //    // cs파일 생성
            //    ClassMaker maker = new ClassMaker(FileManagerList[i].NewFilePath, FileManagerList[i].NewFileName);
            //    maker.AddField(ExcelReader.InfoList[i].DataNames, ExcelReader.InfoList[i].DataTypeCodes);
            //    maker.GenerateCSharpCode();
            //}


            FileManagerList.Clear();
            ExcelReader.Free();

            if(ErrorManager.instance.ErrorLogs.Count > 0)
            {
                ErrorManager.instance.Show();
            }

            ResultTextBox.Text = "변환이 완료되었습니다!!!";
            Process.Start(saveTargetDirectory);
        }


        private void Btn_FileSelected_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = false; // true : 폴더 선택 / false : 파일 선택
            dialog.Filters.Add(new CommonFileDialogFilter("Excel 파일", "*.xlsx")); // 필터 추가

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSysMsg.Text = dialog.FileName;
                currentFileFullPath = dialog.FileName;
                currentDirectory = Path.GetDirectoryName(currentFileFullPath) + "\\";
                currentFileName = Path.GetFileNameWithoutExtension(currentFileFullPath);
            }
        }
        

        private void metroButton1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // true : 폴더 선택 / false : 파일 선택

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                saveTargetDirectory = dialog.FileName +"\\";
            }
        }

        private void CreateSettingsFile(string fileName)
        {
            // 설정값 예시
            //string[] defaultSettings = { "Setting1=Value1", "Setting2=Value2", "Setting3=Value3" };
            string[] defaultSettings = { "" };

            // 파일에 기본 설정값 작성
            File.WriteAllLines(fileName, defaultSettings);
        }

        private static object GetDefaultValue(string typeName)
        {
            switch (typeName.ToLower())
            {
                case "int":
                    return default(int);
                case "float":
                    return default(float);
                case "double":
                    return default(double);
                case "bool":
                    return default(bool);
                case "string":
                    return "";
                default:
                    return null;
            }
        }

        private static void SetArrayData(JArray dataList, string typeName, dynamic value)
        {
            if (value == "" || value == null)
                return;
            
            var dataArray = value.Split(',');
            foreach (var data in dataArray)
            {
                switch (typeName.ToLower())
                {
                    case "intarray":
                        dataList.Add(int.Parse(data));
                        break;
                    case "floatarray":
                        dataList.Add(float.Parse(data));
                        break;
                    case "doublearray":
                        dataList.Add(double.Parse(data));
                        break;
                    case "boolarray":
                        dataList.Add(bool.Parse(data));
                        break;
                    case "stringarray":
                        dataList.Add(data);
                        break;
                }
            }
        }

        private static bool IsArray(string typeName)
        {
            switch (typeName.ToLower())
            {
                case "intarray":
                case "floatarray":
                case "doublearray":
                case "boolarray":
                case "stringarray":
                    return true;
                default:
                    return false;
            }
        }
    }
}
