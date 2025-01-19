using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SLG_ExcelToJson
{
    public class SaveManager
    {
        private string _saveTargetDirectory;
        private List<FileManager> FileManagerList;

        public SaveManager()
        {
            FileManagerList = new List<FileManager>();
        }
        public void Init(string saveTargetDirectory)
        {
            _saveTargetDirectory = saveTargetDirectory;
        }
        public void Save(List<ExcelSheetInfo> infoList , bool isMulti)
        {
            if (isMulti)
            {
                // json 여러개로 뽑을때 사용
                // 엑셀파일 저장.

                foreach (var info in infoList)
                {
                    var jArray = new JArray();
                    var jObj = new JObject();

                    // 데이터 타입 ex(int , string)
                    jObj = ChangeToJObject(info.DataNames, info.DataTypeNames);
                    jArray.Add(jObj);

                    // 데이터 값 ex(1, "홍길동")
                    foreach (var values in info.DataValues)
                    {
                        var jobj = ChangeToJObject(info.DataNames, values, info.DataTypeNames);
                        if (jobj != null)
                            jArray.Add(jobj);
                    }

                    var json = jArray.ToString();
                    var fileName = $"{info.ExcelSheet.Name}Data.json";
                    var filePath = Path.Combine(_saveTargetDirectory, $"{fileName}");
                    var fileManager = new FileManager(filePath);
                    fileManager.FileName = fileName;
                    fileManager.NewFileName =fileName;
                    fileManager.SaveNewFile(json);
                    FileManagerList.Add(fileManager);
                    // cs파일 생성
                    // ClassMaker maker = new ClassMaker(FileManagerList[i].NewFilePath, FileManagerList[i].NewFileName);
                    // maker.AddField(ExcelReader.InfoList[i].DataNames, ExcelReader.InfoList[i].DataTypeCodes);
                    // maker.GenerateCSharpCode();
                }

                var allSheetsValues = ExcelReader.GetAllSheetValues();
                for (int i = 0; i < allSheetsValues.Count; i++)
                {
                    /*
                    var json = string.Empty;
                    var sheetText = JsonChanger.ChangToJArrayToString(ExcelReader.InfoList[i].DataNames, allSheetsValues[i]);

                    var sheetName = ExcelReader.InfoList[i].ExcelSheet.Name;
                    var filePath = Path.Combine(_saveTargetDirectory, $"{sheetName}Data.json");
                    var fileManager = new FileManager(filePath);
                    fileManager.FileName = sheetName;
                    fileManager.NewFileName = $"{sheetName}Data.json";
                    fileManager.SaveNewFile(sheetText);
                    FileManagerList.Add(fileManager);
                    */
                    // FileManagerList[i] = fileManager;
                    
                    // FileManagerList[i].SaveNewFile_Temp(sheetText);
                }
                FileManagerList.Clear();
            }
            else
            {
                var dataDic = new Dictionary<string, JArray>();
                foreach (var info in infoList)
                {
                    var jArray = new JArray();
                    var jObj = new JObject();

                    // 데이터 타입 ex(int , string)
                    jObj = ChangeToJObject(info.DataNames, info.DataTypeNames);
                    jArray.Add(jObj);

                    // 데이터 값 ex(1, "홍길동")
                    foreach (var values in info.DataValues)
                    {
                        var jobj = ChangeToJObject(info.DataNames, values, info.DataTypeNames);
                        if (jobj != null)
                            jArray.Add(jobj);
                    }
                    dataDic.Add(info.ExcelSheet.Name, jArray);
                }
                
                var json = JsonConvert.SerializeObject(dataDic);
                File.WriteAllText($"{_saveTargetDirectory + "GameStaticData"}.json", json);
            }
        }
        
        private JObject ChangeToJObject(List<string> nameList, List<dynamic> valList, List<string> typeList)
        {
            if (valList.Count == 0)
                return null;

            JObject obj = new JObject();
            for (int i = 0; i < nameList.Count; i++)
            {
                if (i >= valList.Count)
                    break;

                var value = valList[i];
                var valueType = typeList[i];
                if (value == null)
                    value = GetDefaultValue(valueType);

                if (IsArray(valueType))
                {
                    var dataList = new JArray();
                    SetArrayData(dataList, valueType, value);
                    value = dataList;
                }

                obj.Add(nameList[i], value);
            }
            return obj;
        }

        private JObject ChangeToJObject(List<string> nameList, List<string> valList)
        {
            if (valList.Count == 0)
                return null;

            JObject obj = new JObject();
            for (int i = 0; i < nameList.Count; i++)
            {
                if (i >= valList.Count)
                    break;
                obj.Add(nameList[i], valList[i]);
            }
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