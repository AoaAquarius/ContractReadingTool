using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinForm1.Models
{
    public class FileModel
    {
        public string FileName { get; set; }
        public IList<string> Paragraphs { get; set; }
        public FileModel()
        {
            Paragraphs = new List<string>();
        }
    }
}
