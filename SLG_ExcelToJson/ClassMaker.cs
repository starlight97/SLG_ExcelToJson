using System;
using System.Collections.Generic;
using System.CodeDom.Compiler;
using System.CodeDom;
using System.Reflection;
using System.IO;

namespace SLG_ExcelToJson
{
    public class ClassMaker
    {
        public string filePath;
        private string fileName;
        public string fileFullName;
        public CodeCompileUnit TargetUnit;
        public CodeTypeDeclaration TargetClass;
        public CodeNamespace TargetNameSpace;

        public ClassMaker(string filePath, string fileName, string nameSpace)
        {
            this.filePath = filePath;
            this.fileName = fileName;
            this.SetFileName();

            TargetUnit = new CodeCompileUnit();
            TargetNameSpace = new CodeNamespace(nameSpace);
            string className = this.fileName;
            className = className.Substring(0, className.LastIndexOf('.'));
            TargetClass = new CodeTypeDeclaration(className);
            //TargetNameSpace.Imports.Add((new CodeNamespaceImport("System")));
            //TargetNameSpace.Imports.Add((new CodeNamespaceImport("System.Collections.Generic")));
            //TargetNameSpace.Imports.Add((new CodeNamespaceImport("System.Linq")));
            //TargetNameSpace.Imports.Add((new CodeNamespaceImport("System.Text")));
            TargetClass.IsClass = true;
            TargetClass.TypeAttributes = TypeAttributes.Public;

            TargetNameSpace.Types.Add(TargetClass);
            TargetUnit.Namespaces.Add(TargetNameSpace);
        }
        public ClassMaker(string filePath, string fileName)
        {
            this.filePath = filePath;
            this.fileName = fileName;
            this.SetFileName();

            TargetUnit = new CodeCompileUnit();
            TargetNameSpace = new CodeNamespace("");
            string className = this.fileName;
            className = className.Substring(0, className.LastIndexOf('.'));
            TargetClass = new CodeTypeDeclaration(className);
            TargetClass.IsClass = true;
            TargetClass.TypeAttributes = TypeAttributes.Public;

            TargetNameSpace.Types.Add(TargetClass);
            TargetUnit.Namespaces.Add(TargetNameSpace);
        }


        private void SetFileName()
        {
            fileName = fileName[0].ToString().ToUpper() + fileName.Substring(1);
            //int underbarIndex = fileName.IndexOf("_");            
            //if(underbarIndex != -1)
            //{
            //    fileName = fileName.Substring(0, underbarIndex) + fileName.Substring(underbarIndex + 1, 1).ToUpper() + fileName.Substring(underbarIndex + 2);
            //}
            fileName = fileName.Replace("_", "");
            fileName = fileName.Replace("data.", "Data.");
            fileName = fileName.Replace(".json", ".cs");
            fileFullName = filePath + "\\"+ fileName;
        }

        public void AddField(List<string> fieldName, List<TypeCode> typeCodes)
        {
            for (int i = 0; i < typeCodes.Count; i++)
            {
                CodeMemberField field = new CodeMemberField();
                field.Attributes = MemberAttributes.Public;
                field.Name = fieldName[i];
                field.Type = new CodeTypeReference(DataTypeChanger.TypeCodeToType(typeCodes[i]));
                this.TargetClass.Members.Add(field);
            }
        }

        public void GenerateCSharpCode()
        {
            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            CodeGeneratorOptions option = new CodeGeneratorOptions();
            option.BracingStyle = "C";
            using (StreamWriter sourceWriter = new StreamWriter(fileFullName))
            {
                provider.GenerateCodeFromCompileUnit(
                    TargetUnit, sourceWriter, option);
            }
        }

    }
}
