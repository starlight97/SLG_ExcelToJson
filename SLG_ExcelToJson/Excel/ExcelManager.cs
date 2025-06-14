using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace SLG_ExcelToJson
{
    public class ExcelManager
    {
        private readonly ExcelReader _excelReader;
        
        private string _settingsFilePath;
        private string _gameDataDirPath;
        private List<string> _targetExcelFileList;

        public ExcelManager()
        {
            _excelReader = new ExcelReader();
        }

        public void Init(string gameDataDirPath)
        {
            _gameDataDirPath = gameDataDirPath;
            
            // GameData 디렉토리가 없으면 생성
            if (!Directory.Exists(_gameDataDirPath))
            {
                Directory.CreateDirectory(_gameDataDirPath);
            }

            _settingsFilePath = Path.Combine(_gameDataDirPath, "setting.txt");

            // settings.txt 파일이 없다면 생성
            if (!File.Exists(_settingsFilePath))
            {
                File.WriteAllText(_settingsFilePath, "# 처리할 엑셀 파일 이름을 한 줄에 하나씩 입력하세요\r\n# 예시:\r\n# Character.xlsx\r\n# Item.xlsx");
            }

            LoadSettings();
            ExcelReader.Init();
        }
        
        public void Clear()
        {
            ExcelReader.Clear();
        }
        
        public List<string> GetTargetExcelFiles()
        {
            var result = new List<string>();
        
            foreach (var fileName in _targetExcelFileList)
            {
                var fullPath = Path.Combine(_gameDataDirPath, fileName);
                if (File.Exists(fullPath))
                {
                    result.Add(fullPath);
                }
            }

            return result;
        }



        /// <summary>
        /// 단일 엑셀 파일을 처리합니다.
        /// </summary>
        public void ProcessSingleFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"파일을 찾을 수 없습니다: {filePath}");

            ExcelReader.AddExcelFile(filePath);
        }
        

        /// <summary>
        /// 모든 시트의 데이터를 가져옵니다.
        /// </summary>
        public List<List<List<dynamic>>> GetAllSheetValues()
        {
            return ExcelReader.GetAllSheetValues();
        }

        /// <summary>
        /// 특정 시트의 데이터를 가져옵니다.
        /// </summary>
        public List<List<dynamic>> GetSheetValuesByIndex(int index)
        {
            return ExcelReader.GetSheetValuesByIndex(index);
        }

        /// <summary>
        /// 현재 처리된 엑셀 파일의 정보 목록을 반환합니다.
        /// </summary>
        public List<ExcelSheetInfo> GetInfoList()
        {
            return ExcelReader.InfoList;
        }
    
        
        
        private void LoadSettings()
        {
            // settings.txt 파일이 없다면 생성
            if (!File.Exists(_settingsFilePath))
            {
                File.WriteAllText(_settingsFilePath, "# 처리할 엑셀 파일 이름을 한 줄에 하나씩 입력하세요\r\n# 예시:\r\n# Character.xlsx\r\n# Item.xlsx");
            }

            // settings.txt 파일에서 엑셀 파일 목록 읽기
            _targetExcelFileList = File.ReadAllLines(_settingsFilePath)
                .Where(line => !string.IsNullOrWhiteSpace(line) && !line.TrimStart().StartsWith("#"))
                .Select(line => line.Trim())
                .ToList();
        }

    }
}