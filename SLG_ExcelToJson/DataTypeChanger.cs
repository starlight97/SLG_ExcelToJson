using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.ComponentModel;

namespace SLG_ExcelToJson
{
    public static class DataTypeChanger
    {
        public enum eDataType
        {
            [Description("int")]
            Int = TypeCode.Int32,
            [Description("long")]
            Int64 = TypeCode.Int64,
            [Description("float")]
            Single = TypeCode.Single,
            [Description("double")]
            Double = TypeCode.Double,
            [Description("char")]
            Char = TypeCode.Char,
            [Description("string")]
            String = TypeCode.String,
            [Description("bool")]
            Boolean = TypeCode.Boolean,
            [Description("DateTime")]
            DateTime = TypeCode.DateTime
        }

        public static Type TypeCodeToType(TypeCode code)
        {
            switch (code)
            {
                case TypeCode.Boolean:
                    return typeof(bool);

                case TypeCode.Byte:
                    return typeof(byte);

                case TypeCode.Char:
                    return typeof(char);

                case TypeCode.DateTime:
                    return typeof(DateTime);

                case TypeCode.DBNull:
                    return typeof(DBNull);

                case TypeCode.Decimal:
                    return typeof(decimal);

                case TypeCode.Double:
                    return typeof(double);

                case TypeCode.Empty:
                    return null;

                case TypeCode.Int16:
                    return typeof(short);

                case TypeCode.Int32:
                    return typeof(int);

                case TypeCode.Int64:
                    return typeof(long);

                case TypeCode.Object:
                    return typeof(object);

                case TypeCode.SByte:
                    return typeof(sbyte);

                case TypeCode.Single:
                    return typeof(Single);

                case TypeCode.String:
                    return typeof(string);

                case TypeCode.UInt16:
                    return typeof(UInt16);

                case TypeCode.UInt32:
                    return typeof(UInt32);

                case TypeCode.UInt64:
                    return typeof(UInt64);

                default:
                    return typeof(string);
            }
        }

        public static TypeCode GetTypeCodeByDescription(string desc)
        {
        //eDataType의 값을 모두 돌며 Descript값 확인.
        foreach (var eVal in typeof(eDataType).GetEnumValues())
        {
            FieldInfo field = typeof(eDataType).GetField(eVal.ToString());
            DescriptionAttribute att = (DescriptionAttribute)field.GetCustomAttribute(typeof(DescriptionAttribute), false);
            if (att.Description == desc)
                return (TypeCode)eVal;
        }
        //없다면 모두 String으로 보겠음.
        return TypeCode.String;
        }

        public static dynamic GetValue(TypeCode typeCode, dynamic value)
        {
            if (typeCode == TypeCode.DateTime)
                return ParseDateTime(value);

            else
            {
                try
                {
                    return Convert.ChangeType(value, typeCode);
                }
                catch (Exception e)
                {
                    ErrorManager.instance.AddErrorLog($"ERROR : {value} {typeCode}");
                    return null;
                }

                
            }
        }

        public static DateTime ParseDateTime(dynamic dateTime)
        {
            return DateTime.Parse(dateTime.ToString());
        }
    }
}
