using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;

namespace SLG_ExcelToJson
{
    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        private const int DEFAULT_WIDTH = 382;
        private const int EXPANDED_WIDTH = 1920;
        
        private const int DEFAULT_HEIGHT = 351;
        private const int EXPANDED_HEIGHT = 1080;
        

        private ExcelManager _excelManager;
        private SaveManager _saveManager;
        private List<SLGFile> _fileList;
        private string _currentFileFullPath;
        private string _gameDataDirPath;

        private string _saveTargetDirectory = string.Empty;
        private bool _useAutoSet;
        private bool _isExpanded = false;


        public MainForm()
        {
            _fileList = new List<SLGFile>();
            _saveManager = new SaveManager();
            _excelManager = new ExcelManager();

            InitializeComponent();
            
            Height = DEFAULT_HEIGHT;
            Width = DEFAULT_WIDTH;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Init();
        }
        
        private void Init()
        {
            _useAutoSet = Chk_UseAutoSet.Checked;

            if (_useAutoSet)
            {
                // 현재 프로그램 실행 경로
                var currentPath = Directory.GetCurrentDirectory();
                var gameDataPath = Path.GetFullPath(Path.Combine(currentPath, @"..\..\..\GameData\GameStaticData.xlsx"));
                var gameDataDirPath = Path.GetFullPath(Path.Combine(currentPath, @"..\..\..\GameData\"));
                var jsonExportPath = Path.GetFullPath(Path.Combine(currentPath, @"..\..\Assets\Resources\Datas"));
                
                _currentFileFullPath = Path.GetFullPath(gameDataPath);
                _gameDataDirPath = Path.GetFullPath(gameDataDirPath);
                // _saveTargetDirectory = Path.GetFullPath(jsonExportPath);
                _saveTargetDirectory = Path.GetFullPath(gameDataDirPath);

                AddDebugLog(_gameDataDirPath);
            }
            else
            {
                var fileName = "settings.txt";
                // 파일이 존재하는지 확인
                if (!File.Exists(fileName))
                {
                    // 파일이 없는 경우에는 새로운 파일을 생성
                    CreateSettingsFile(fileName);
                }
                var settingValue = File.ReadAllLines("settings.txt");


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
        
        private void CreateSettingsFile(string fileName)
        {
            // 설정값 예시
            //string[] defaultSettings = { "Setting1=Value1", "Setting2=Value2", "Setting3=Value3" };
            string[] defaultSettings = { "" };

            // 파일에 기본 설정값 작성
            File.WriteAllLines(fileName, defaultSettings);
        }

        private bool IsValid(out string errorMsg)
        {
            errorMsg = string.Empty;
            // if (_currentFileFullPath == null || File.Exists(_currentFileFullPath) == false)
            // {
            //     errorMsg = $"{_currentFileFullPath} 변환할 파일이 없습니다.";
            //     return false;
            // }
            
            return true;
        }

        private void AddDebugLog(string log)
        {
            TXB_DebugLog.Text += $"\r\n{log}";
        }
        
        
        #region OnClick

        private void OnClickConvert(object sender, EventArgs e)
        {
            _excelManager.Init(_gameDataDirPath);
            var dataFilePathList = _excelManager.GetTargetExcelFiles();
            if (dataFilePathList.Count == 0)
            {
                MessageBox.Show("변환할 파일이 없습니다",
                                "Error",
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Information, 
                                MessageBoxDefaultButton.Button2);
                BtnSysLog.Text = "변환 준비중...";
                return;
            }
            
            var isValid = IsValid(out var errorMsg);
            if (isValid == false)
            {
                MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, 
                    MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                BtnSysLog.Text = "변환 준비중...";
                return;
            }
            try
            {
                foreach (var dataFilePath in dataFilePathList)
                {
                    _excelManager.ProcessSingleFile(dataFilePath);
                }
                
                _saveManager.Init(_saveTargetDirectory);
                _saveManager.Save(_excelManager.GetInfoList(), true);
                
                ErrorManager.instance.Show();
                _fileList.Clear();
                
                Process.Start(_saveTargetDirectory);
                ErrorManager.instance.Clear();
                BtnSysLog.Text = "변환이 완료되었습니다!!!";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"변환 중 오류가 발생했습니다: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _excelManager.Clear();
            }
        }

        private void OnClickConvertLegacy(object sender, EventArgs e)
        {
            //if (Directory.Exists(currentDirectory + "\\json") == false)
            //{
            //    Directory.CreateDirectory(currentDirectory + "\\json");
            //}
            //if (Directory.Exists(currentDirectory + "\\cs") == false)
            //{
            //    Directory.CreateDirectory(currentDirectory + "\\cs");
            //}

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
        }
        
        private void OnClickClose(object sender, EventArgs e)
        {
            string targetFilePath = _currentFileFullPath;
            string targetSaveDirectoryPath = _saveTargetDirectory;

            string settingValue = targetFilePath + "\r\n";
            settingValue += targetSaveDirectoryPath + "\r\n";
            File.WriteAllText($"settings.txt", settingValue);

            Application.Exit();
        }

        private void OnClickDirectoryOpen(object sender, EventArgs e)
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

        private void OnClickFileSelected(object sender, EventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = false; // true : 폴더 선택 / false : 파일 선택
            dialog.Filters.Add(new CommonFileDialogFilter("Excel 파일", "*.xlsx")); // 필터 추가

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtSysMsg.Text = dialog.FileName;
                _currentFileFullPath = dialog.FileName;
            }
        }

        private void OnClickSaveDirectoryOpen(object sender, EventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // true : 폴더 선택 / false : 파일 선택

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                _saveTargetDirectory = dialog.FileName +"\\";
            }
        }

        private void OnClickDebugLog(object sender, EventArgs e)
        {
            return;
            
            _isExpanded = !_isExpanded;
            
            Height = _isExpanded ? EXPANDED_HEIGHT : DEFAULT_HEIGHT;
            Width = _isExpanded ? EXPANDED_WIDTH : DEFAULT_WIDTH;
        }
        
        #endregion
    }
}
