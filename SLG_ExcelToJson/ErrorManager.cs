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
        public int ErrorLogCount => _errorLogList.Count;

        private List<string> _errorLogList = new List<string>();
        
        public void Clear()
        {
            _errorLogList.Clear();
        }

        public void AddErrorLog(string errorLog)
        {
            _errorLogList.Add(errorLog);
        }

        public void Show()
        {
            if (_errorLogList.Count == 0)
            {
                return;
            }
            
            var stringBuilder = new StringBuilder();
            foreach (var errorlog in _errorLogList)
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
