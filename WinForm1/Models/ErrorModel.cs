using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinForm1.Models
{
    public class ErrorModel
    {
        public string FileName { get; set; }
        public string ErrorMessage { get; set; }
        public ErrorModel(string fileName, string message)
        {
            this.FileName = fileName;
            this.ErrorMessage = message;
        }
    }
}
