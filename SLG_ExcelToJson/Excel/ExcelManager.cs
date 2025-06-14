using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace SLG_ExcelToJson
{
    public class ExcelManager
    {
        private readonly string[] EXCEL_EXTENSIONS = { ".xlsx", ".xls" };
        private List<string> _excelFiles = new List<string>();
        private List<string> _filteredFiles = new List<string>();
        private string _selectedFolderPath;
        
        
        /// <summary>
        /// 사용자가 폴더를 선택하고 엑셀 파일들을 읽어오는 메서드
        /// </summary>
        /// <returns>선택된 폴더 경로</returns>
        public string SelectFolder()
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.Title = "엑셀 파일이 있는 폴더를 선택하세요";

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    _selectedFolderPath = dialog.FileName;
                    LoadExcelFiles();
                    return _selectedFolderPath;
                }
                
                return string.Empty;
            }
        }

        /// <summary>
        /// 선택된 폴더에서 엑셀 파일들을 읽어오는 메서드
        /// </summary>
        private void LoadExcelFiles()
        {
            _excelFiles.Clear();
            
            try
            {
                var files = Directory.GetFiles(_selectedFolderPath)
                                   .Where(file => EXCEL_EXTENSIONS.Contains(Path.GetExtension(file).ToLower()));
                
                _excelFiles.AddRange(files);
            }
            catch (Exception ex)
            {
                throw new Exception($"엑셀 파일 로딩 중 오류 발생: {ex.Message}");
            }
        }

        /// <summary>
        /// 필터 텍스트 파일을 선택하고 필터링을 수행하는 메서드
        /// </summary>
        /// <returns>필터링된 파일 목록</returns>
        public List<string> FilterExcelFiles()
        {
            using (CommonOpenFileDialog dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = false;
                dialog.Title = "필터 텍스트 파일을 선택하세요";
                dialog.Filters.Add(new CommonFileDialogFilter("Text files", "*.txt"));

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    return ApplyFilter(dialog.FileName);
                }
                
                return new List<string>();
            }
        }

        /// <summary>
        /// 텍스트 파일 내용을 기반으로 엑셀 파일을 필터링하는 메서드
        /// </summary>
        /// <param name="filterFilePath">필터 텍스트 파일 경로</param>
        /// <returns>필터링된 파일 목록</returns>
        private List<string> ApplyFilter(string filterFilePath)
        {
            _filteredFiles.Clear();
            
            try
            {
                // 텍스트 파일에서 파일명 목록 읽기
                var filterNames = File.ReadAllLines(filterFilePath)
                                    .Where(line => !string.IsNullOrWhiteSpace(line))
                                    .Select(line => line.Trim())
                                    .ToList();

                // 파일명과 확장자를 분리하여 비교
                foreach (var excelFile in _excelFiles)
                {
                    var fileName = Path.GetFileNameWithoutExtension(excelFile);
                    if (filterNames.Contains(fileName))
                    {
                        _filteredFiles.Add(excelFile);
                    }
                }

                return _filteredFiles;
            }
            catch (Exception ex)
            {
                throw new Exception($"파일 필터링 중 오류 발생: {ex.Message}");
            }
        }

        /// <summary>
        /// 필터링된 파일 목록을 반환하는 메서드
        /// </summary>
        public List<string> GetFilteredFiles()
        {
            return _filteredFiles;
        }
    }
}