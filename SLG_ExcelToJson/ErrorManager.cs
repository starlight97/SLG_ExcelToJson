using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SLG_ExcelToJson
{
    public sealed class ErrorManager
    {
        public static readonly ErrorManager instance = new ErrorManager();

        public List<string> ErrorLogs = new List<string>();

        public void Init()
        {
            ErrorLogs.Clear();
        }

        public void AddErrorLog(string errorLog)
        {
            ErrorLogs.Add(errorLog);
        }

        public void Show()
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (var errorlog in ErrorLogs)
            {
                stringBuilder.AppendLine(errorlog);
            }

            MessageBox.Show(stringBuilder.ToString());
        }

        private ErrorManager()
        {
        }
    }


}
