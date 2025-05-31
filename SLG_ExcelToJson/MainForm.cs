using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Reflection;

namespace SLG_ExcelToJson
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        public List<FileManager> FileManagerList;


        private SaveManager _saveManager;
        private List<string> _excelPathList;
        private string _currentFileFullPath;

        private string _saveTargetDirectory = string.Empty;
        private bool _useAutoSet;

        public MainForm()
        {
            FileManagerList = new List<FileManager>();
            _saveManager = new SaveManager();

            _excelPathList = new List<string>();

            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Init();
        }

        private void mbtClose_Click(object sender, EventArgs e)
        {
            string targetFilePath = _currentFileFullPath;
            string targetSaveDirectoryPath = _saveTargetDirectory;

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

            if (_currentFileFullPath == null || File.Exists(_currentFileFullPath) == false)
            {
                MessageBox.Show($"{_currentFileFullPath} 변환할 파일이 없습니다."
                                , "아이고..."
                                , MessageBoxButtons.OK,MessageBoxIcon.Information
                                , MessageBoxDefaultButton.Button2);
                ResultTextBox.Text = "변환 준비중...";
                return;
            }

            ExcelReader.Init();
            ExcelReader.AddExcelFile(_currentFileFullPath);
            
            _saveManager.Init(_saveTargetDirectory);
            _saveManager.Save(ExcelReader.InfoList, true);
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
            Process.Start(_saveTargetDirectory);
        }


        private void Btn_FileSelected_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = false; // true : 폴더 선택 / false : 파일 선택
            dialog.Filters.Add(new CommonFileDialogFilter("Excel 파일", "*.xlsx")); // 필터 추가

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSysMsg.Text = dialog.FileName;
                _currentFileFullPath = dialog.FileName;
            }
        }
        

        private void metroButton1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // true : 폴더 선택 / false : 파일 선택

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                _saveTargetDirectory = dialog.FileName +"\\";
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

        private void Init()
        {
            _useAutoSet = Chk_UseAutoSet.Checked;

            if (_useAutoSet)
            {
                // 현재 프로그램 실행 경로
                string currentPath = Directory.GetCurrentDirectory();
                string gameDataPath = Path.GetFullPath(Path.Combine(currentPath, @"..\..\..\GameData\GameStaticData.xlsx"));
                string jsonExportPath = Path.GetFullPath(Path.Combine(currentPath, @"..\..\Assets\Resources\Datas"));
                
                _currentFileFullPath = Path.GetFullPath(gameDataPath);
                _saveTargetDirectory = Path.GetFullPath(jsonExportPath); 
            }
            else
            {
                string fileName = "settings.txt";
                // 파일이 존재하는지 확인
                if (!File.Exists(fileName))
                {
                    // 파일이 없는 경우에는 새로운 파일을 생성
                    CreateSettingsFile(fileName);
                }
                string[] settingValue = File.ReadAllLines("settings.txt");


                for (int index = 0; index < settingValue.Count(); index++)
                {
                    if (settingValue[index] == string.Empty)
                        continue;

                    if(index == 0)
                    {
                        _currentFileFullPath = Path.GetFullPath(settingValue[0]); // 첫 번째 줄은 currentFileFullPath
                    }

                    if(index == 1)
                    {
                        _saveTargetDirectory = Path.GetFullPath(settingValue[1]); // 두 번째 줄은 saveTargetDirectory
                    }
                }   
            }
            
            txtSysMsg.Text = _currentFileFullPath;
        }

        private void Chk_UseAutoSet_CheckedChanged(object sender, EventArgs e)
        {
            _useAutoSet = Chk_UseAutoSet.Checked;
        }
    }
}
