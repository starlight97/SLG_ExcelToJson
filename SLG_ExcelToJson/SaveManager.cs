using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SLG_ExcelToJson
{
    public class SaveManager
    {
        private string _saveTargetDirectory;
        private List<SLGFile> FileManagerList;

        public SaveManager()
        {
            FileManagerList = new List<SLGFile>();
        }
        public void Init(string saveTargetDirectory)
        {
            _saveTargetDirectory = saveTargetDirectory;
        }
        
        
        public bool Save(List<ExcelSheetInfo> infoList , bool isMulti)
        {
            if (isMulti)
            {
                // json 여러개로 뽑을때 사용
                // 엑셀파일 저장.

                foreach (var info in infoList)
                {
                    var excelFileName = $"{info.FileName}.xlsx";
                    
                    var jArray = new JArray();
                    var jObj = new JObject();

                    // 데이터 타입 ex(int , string)
                    jObj = ChangeToJObject(out var successType, info.DataNameList, info.DataTypeNameList);
                    if (successType == false)
                    {
                        var errorMsg = $"데이터 변환 오류 : ExcelFileName : {excelFileName}";
                        ErrorManager.instance.AddErrorLog(errorMsg);
                        return false;
                    }
                    
                    jArray.Add(jObj);
                    // 데이터 값 ex(1, "홍길동")
                    foreach (var values in info.DataValues)
                    {
                        var jobj = ChangeToJObject(out var successValue, info.DataNameList, values, info.DataTypeNameList);
                        if (successValue == false)
                        {
                            var errorMsg = $"데이터 변환 오류 : ExcelFileName : {excelFileName}";
                            ErrorManager.instance.AddErrorLog(errorMsg);
                            return false;
                        }
                        
                        jArray.Add(jobj);
                    }

                    var json = jArray.ToString();
                    var saveFileName = $"{info.ExcelSheet.Name}Data.json";
                    var filePath = Path.Combine(_saveTargetDirectory, $"{saveFileName}");
                    var fileManager = new SLGFile(filePath);
                    fileManager.FileName = saveFileName;
                    fileManager.NewFileName = saveFileName;
                    fileManager.SaveNewFile(json);
                    // cs파일 생성
                    // ClassMaker maker = new ClassMaker(FileManagerList[i].NewFilePath, FileManagerList[i].NewFileName);
                    // maker.AddField(ExcelReader.InfoList[i].DataNames, ExcelReader.InfoList[i].DataTypeCodes);
                    // maker.GenerateCSharpCode();
                }
            }
            else
            {
                var dataDic = new Dictionary<string, JArray>();
                foreach (var info in infoList)
                {
                    var jArray = new JArray();
                    var jObj = new JObject();

                    // 데이터 타입 ex(int , string)
                    jObj = ChangeToJObject(out var successType, info.DataNameList, info.DataTypeNameList);
                    jArray.Add(jObj);

                    // 데이터 값 ex(1, "홍길동")
                    foreach (var values in info.DataValues)
                    {
                        var jobj = ChangeToJObject(out var successValue, info.DataNameList, values, info.DataTypeNameList);
                        if (jobj != null)
                            jArray.Add(jobj);
                    }
                    dataDic.Add(info.ExcelSheet.Name, jArray);
                }
                
                var json = JsonConvert.SerializeObject(dataDic);
                File.WriteAllText($"{_saveTargetDirectory + "GameStaticData"}.json", json);
            }

            return true;
        }
        
        private JObject ChangeToJObject(out bool success, List<string> nameList, List<dynamic> valList, List<string> typeList)
        {
            success = false;
            if (valList.Count == 0)
            {
                return null;
            }

            var obj = new JObject();
            var count = Math.Min(nameList.Count, valList.Count);
            for (int i = 0; i < count; i++)
            {
                var value = valList[i];
                var valueType = typeList[i];
                if (value == null)
                {
                    value = GetDefaultValue(valueType);
                }

                if (IsArray(valueType))
                {
                    var dataList = new JArray();
                    SetArrayData(dataList, valueType, value);
                    value = dataList;
                }

                obj.Add(nameList[i], value);
            }

            success = true;
            return obj;
        }

        private JObject ChangeToJObject(out bool success, List<string> nameList, List<string> valList)
        {
            success = false;
            if (valList.Count == 0)
            {
                return null;
            }

            var obj = new JObject();
            var count = Math.Min(nameList.Count, valList.Count);
            for (int i = 0; i < count; i++)
            {
                try
                {
                    obj.Add(nameList[i], valList[i]);
                }
                catch (Exception ex)
                {
                    var nameValue = nameList[i] == null ? "null" : nameList[i];
                    var valValue = valList[i] == null ? "null" : valList[i];
                    var errorMsg = $"Index: {i}, Name: {nameValue}, Value: {valValue} \r\n" 
                                   + $"Error: {ex.Message}";
                    ErrorManager.instance.AddErrorLog(errorMsg);
                    return null;
                }
            }
            success = true;
            return obj;
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