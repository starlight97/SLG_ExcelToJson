using System;
using System.Linq;
using System.IO;

namespace SLG_ExcelToJson
{
    public class SLGFile
    {
        //파일의 경로만.
        private string @filePath;
        private string @fileName;
        private string @originFileName;

        public string @FilePath
        {
            get { return this.filePath; }
            set
            {
                //끝에 \가 포함되어 있다면 제거.
                filePath = (value.Last() == '\\' || value.Last() == '/') ?
                    value.Substring(0, value.Length - 1) : value;

                //새로 저장될 패쓰를 자동으로 설정.
                if (string.IsNullOrEmpty(newFilePath))
                    NewFilePath = filePath;
            }
        }
        public string @FileName
        {
            get { return fileName; }
            set
            {
                fileName = value;
                if (string.IsNullOrEmpty(newFileName))
                    NewFileName = value.Contains('.') ? value.Substring(0, value.LastIndexOf('.')) : value;
            }
        }
        public string @OriginFileName
        {
            get { return originFileName; }
        }

        public string @FileFullPath
        {
            get { return filePath + '\\' + fileName; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    //FilePath와 FileName 자동으로 등록.
                    int lastIndex = value.LastIndexOf('\\');
                    FilePath = value.Substring(0, lastIndex);
                    FileName = value.Substring(lastIndex + 1);
                }
            }
        }
        public string @FileExtension
        {
            get
            {
                if (!string.IsNullOrEmpty(this.fileName) && this.fileName.Contains("."))
                    return this.fileName.Substring(this.fileName.LastIndexOf('.') + 1);

                return "";
            }

            set
            {
                //파일이름이 있을 때.
                if (!string.IsNullOrEmpty(this.fileName))
                {
                    FileName = FileName.Substring(0, fileName.LastIndexOf('.')) +
                        ((value.Contains('.')) ? value : "." + value);
                }
            }
        }

        private string @newFileName;
        private string @newFilePath;

        public string @NewFileName
        {
            get { return this.newFileName; }
            set
            {
                //파일 이름 없을 때.
                if (string.IsNullOrEmpty(newFileName))
                {
                    if (value.Contains('.'))
                        newFileName = value;
                    else
                        newFileName = value + ".txt";
                }

                //이미 있을 때. 확장자만 가져옴.
                else
                {
                    //확장자를 가지고 있다면 그냥 넣겠음.
                    if (value.Contains('.'))
                        newFileName = value;
                    else
                    {
                        string extension = newFileName.Substring(newFileName.LastIndexOf('.') + 1);
                        newFileName = value + '.' + extension;
                    }
                }
            }
        }
        public string @NewFilePath
        {
            get { return newFilePath; }
            set { newFilePath = (value.Last() == '\\' || value.Last() == '/') ? value.Substring(0, value.Length - 1) : value; }
        }
        public string @NewFileFullPath
        {
            get { return newFilePath + @"\" + newFileName; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    newFileName = value.Substring(value.LastIndexOf('\\') + 1);
                    newFilePath = value.Substring(0, value.LastIndexOf('\\'));
                }
            }
        }
        public string @NewFileExtension
        {
            get
            {
                if (!string.IsNullOrEmpty(newFileName) && newFileName.Contains('.'))
                    return newFileName.Substring(newFileName.LastIndexOf('.') + 1);

                return "";
            }
            set
            {
                string extension = (value.Contains('.')) ? value.Substring(1) : value;
                if (!string.IsNullOrEmpty(newFileName))
                    NewFileName = (newFileName.Contains('.')) ?
                        newFileName.Substring(0, newFileName.LastIndexOf('.')) + '.' + extension :
                        newFileName + '.' + extension;
                else
                {
                    newFileName = '.' + extension;
                }
            }
        }

        public SLGFile() { }
        public SLGFile(string fileFullPath, string newFileFullPath = "")
        {
            this.FileFullPath = fileFullPath;
            this.NewFileFullPath = newFileFullPath;
        }


        public void SaveNewFile(string text)
        {
            File.WriteAllText(NewFileFullPath, text);
        }
        // 임시
        public void SaveNewFile_Temp(string text)
        {
            File.WriteAllText(newFilePath + "\\json" + @"\" + newFileName.ToLower(), text);
        }

        public bool NewFileExist()
        {
            return File.Exists(NewFileFullPath);
        }

        public void PrintFileInfo()
        {
            Console.WriteLine("{0}\n{1}\n{2}\n{3}",
                this.FileFullPath, this.FilePath, this.FileName, this.FileExtension);
        }

        public void PrintNewFileInfo()
        {
            Console.WriteLine("{0}\n{1}\n{2}\n{3}",
                this.NewFileName, this.NewFilePath, this.NewFileFullPath, this.NewFileExtension);
        }
    }
}
