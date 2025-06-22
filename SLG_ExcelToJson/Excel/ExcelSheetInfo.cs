using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SLG_ExcelToJson
{
    public class ExcelSheetInfo
    {
        public string FileName => excelSheet?.Parent?.Name ?? "Unknown";
        public List<string> DataNameList => _dataNameList;
        public List<TypeCode> DataTypeCodeList => _dataTypeCodeList;
        public List<string> DataTypeNameList => _dataTypeNameList;
        public List<List<dynamic>> DataValues => _dataValues;
        public Range UsedRange => _usedRange;

        public Worksheet ExcelSheet
        {
            get { return this.excelSheet; }
            set
            {
                excelSheet = value;

                //usedRange 자동 할당.
                _usedRange = excelSheet.UsedRange;

                //데이터 네임들 자동으로 뽑아줌.
                int row = 1;
                for (int col = 1; col <= _usedRange.Columns.Count; col++)
                {
                    if (_usedRange.Cells[row, col] != null && _usedRange.Cells[row, col].Value != null)
                    {
                        colCount++;
                        _dataNameList.Add(_usedRange.Cells[row, col].Value.ToString());
                    }
                }
                //데이터 타입스트링 자동으로 뽑아줌.
                row = 2;
                for (int col = 1; col <= _usedRange.Columns.Count; col++)
                {
                    if (_usedRange.Cells[row, col] != null && _usedRange.Cells[row, col].Value != null)
                    {
                        var typeName = _usedRange.Cells[row, col].Value.ToString();

                        if(typeName == "Int" || typeName == "Long" || typeName == "Float"
                            || typeName == "Double" || typeName == "Char" || typeName == "String"
                            || typeName == "Bool")
                        {
                            typeName = typeName.ToLower();
                        }

                        _dataTypeNameList.Add(typeName);                       
                    }
                }

                //데이터 타입코드 자동으로 뽑아줌.
                foreach (string type in _dataTypeNameList)
                {
                    _dataTypeCodeList.Add(DataTypeChanger.GetTypeCodeByDescription(type));
                    //Console.WriteLine("{0}", DataTypeChanger.GetTypeCodeByDescription(type));
                }

                //밸류들 자동으로 뽑아줌.
                _dataValues = GetSheetValues();
            }
        }
        
        private Range _usedRange;
        private Worksheet excelSheet;
        private List<string> _dataNameList = new List<string>();
        private List<TypeCode> _dataTypeCodeList = new List<TypeCode>();
        private List<string> _dataTypeNameList = new List<string>();
        private List<List<dynamic>> _dataValues = new List<List<dynamic>>();
        private int colCount = 0;
        private int rowCount = 0;
        

        public List<List<dynamic>> GetSheetValues()
        {
            // NULL 시트 체크 ID 값이 비어 있다면 NULL
            bool nullCheck = false;
            List<List<dynamic>> rtnList = new List<List<dynamic>>();
            for (int row = 3; row <= _usedRange.Rows.Count; row++)
            {
                List<dynamic> valList = new List<dynamic>();

                for (int col = 1; col <= colCount; col++)
                {
                    if (_usedRange.Cells[row, col].Value != null)
                    {
                        dynamic value = null;

                        TypeCode type = this._dataTypeCodeList[col - 1];
                        value = DataTypeChanger.GetValue(type, _usedRange.Cells[row, col].Value);
                        //Console.WriteLine("test : value : {0}", value);
                        valList.Add(value);
                    }
                    else
                    {
                        if (col == 1)
                        {
                            nullCheck = true;
                            break;
                        }
                        //Console.WriteLine("{0}", value);
                        valList.Add(null);
                    }
                }
                if (nullCheck)
                    break;

                rtnList.Add(valList);
            }
            return rtnList;
        }

        /// <summary>
        /// _로 시작하는건 미사용 데이터
        /// </summary>
        public void RemoveUnUsedData()
        {
            var columnCount = _dataNameList.Count;
            for (int i = columnCount - 1; i >= 0; --i)
            {
                var columnName = _dataNameList[i];
                var skipColumn = columnName.StartsWith("_");
                if (skipColumn == false)
                {
                    continue;
                }
                
                _dataNameList.RemoveAt(i);
                _dataTypeCodeList.RemoveAt(i);
                _dataTypeNameList.RemoveAt(i);
            
                foreach (var dataValue in _dataValues)
                {
                    dataValue.RemoveAt(i);
                }
            }
        }

        public void PrintDataTypes()
        {
            for (int i = 0; i < _dataTypeCodeList.Count; i++)
            {
                Console.WriteLine(ExcelSheet.Name);
                Console.WriteLine("{0}, {1}, {2}",
                    _dataNameList[i], _dataTypeNameList[i], _dataTypeCodeList[i]);
            }
        }

        public void Clear()
        {
            Marshal.ReleaseComObject(_usedRange);
            Marshal.ReleaseComObject(excelSheet);
        }
    }
}