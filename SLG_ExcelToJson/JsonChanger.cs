using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace SLG_ExcelToJson
{
    public static class JsonChanger
    {
        public static JObject ChangeToJObject(List<string> nameList, List<dynamic> valList)
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

        public static JArray ChangeToJArray(List<string> nameList, List<List<dynamic>> valList)
        {
            JArray rtnArr = new JArray();
            for (int i = 0; i < valList.Count; i++)
            {
                var jobj = JsonChanger.ChangeToJObject(nameList, valList[i]);
                if(jobj != null)
                    rtnArr.Add(jobj);
            }
            return rtnArr;
        }

        public static string ChangToJArrayToString(List<string> nameList, List<List<dynamic>> valList)
        {
            return ChangeToJArray(nameList, valList).ToString();
        }
    }
}
