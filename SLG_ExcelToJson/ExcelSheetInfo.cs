using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SLG_ExcelToJson
{
    public class ExcelSheetInfo
    {
        public List<string> DataNames { get { return this.dataNames; } }
        public List<TypeCode> DataTypeCodes { get { return this.dataTypeCodes; } }
        public List<string> DataTypeNames { get { return this.dataTypeNames; } }
        public List<List<dynamic>> DataValues { get { return this.dataValues; } }
        public Range UsedRange { get { return this.usedRange; } }
        
        public Worksheet ExcelSheet
        {
            get { return this.excelSheet; }
            set
            {
                excelSheet = value;

                //usedRange 자동 할당.
                usedRange = excelSheet.UsedRange;

                //데이터 네임들 자동으로 뽑아줌.
                int row = 1;
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    if (usedRange.Cells[row, col] != null && usedRange.Cells[row, col].Value != null)
                    {
                        colCount++;
                        dataNames.Add(usedRange.Cells[row, col].Value.ToString());
                    }
                }
                //데이터 타입스트링 자동으로 뽑아줌.
                row = 2;
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    if (usedRange.Cells[row, col] != null && usedRange.Cells[row, col].Value != null)
                    {
                        string typeName = usedRange.Cells[row, col].Value.ToString();

                        if(typeName == "Int" || typeName == "Long" || typeName == "Float"
                            || typeName == "Double" || typeName == "Char" || typeName == "String"
                            || typeName == "Bool")
                        {
                            typeName = typeName.ToLower();
                        }

                        dataTypeNames.Add(typeName);                       
                    }
                }

                //데이터 타입코드 자동으로 뽑아줌.
                foreach (string type in dataTypeNames)
                {
                    dataTypeCodes.Add(DataTypeChanger.GetTypeCodeByDescription(type));
                    //Console.WriteLine("{0}", DataTypeChanger.GetTypeCodeByDescription(type));
                }

                //밸류들 자동으로 뽑아줌.
                dataValues = this.GetSheetValues();
            }
        }

        public ExcelSheetInfo()
        {
            this.dataNames = new List<string>();
            this.dataTypeCodes = new List<TypeCode>();
            this.dataTypeNames = new List<string>();
        }
        
        private Range usedRange;
        private Worksheet excelSheet;
        private List<string> dataNames = new List<string>();
        private List<TypeCode> dataTypeCodes = new List<TypeCode>();
        private List<string> dataTypeNames = new List<string>();
        private List<List<dynamic>> dataValues = new List<List<dynamic>>();
        private int colCount = 0;
        private int rowCount = 0;
        

        public List<List<dynamic>> GetSheetValues()
        {
            // NULL 시트 체크 ID 값이 비어 있다면 NULL
            bool nullCheck = false;
            List<List<dynamic>> rtnList = new List<List<dynamic>>();
            for (int row = 3; row <= usedRange.Rows.Count; row++)
            {
                List<dynamic> valList = new List<dynamic>();

                for (int col = 1; col <= colCount; col++)
                {
                    if (usedRange.Cells[row, col].Value != null)
                    {
                        dynamic value = null;

                        TypeCode type = this.dataTypeCodes[col - 1];
                        value = DataTypeChanger.GetValue(type, usedRange.Cells[row, col].Value);
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

        public void PrintDataTypes()
        {
            for (int i = 0; i < this.dataTypeCodes.Count; i++)
            {
                Console.WriteLine(this.ExcelSheet.Name);
                Console.WriteLine("{0}, {1}, {2}",
                    this.dataNames[i], this.dataTypeNames[i], this.dataTypeCodes[i]);
            }
        }

        public void Free()
        {
            Marshal.ReleaseComObject(this.usedRange);
            Marshal.ReleaseComObject(this.excelSheet);
        }
    }
}
