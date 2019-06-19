using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinForm1.Models
{
    public class CompanyModel
    {
        public IList<FileModel> FileModels;
        public IList<string> AmendmentFileNames { get; set; }

        public CompanyModel()
        {
            FileModels = new List<FileModel>();
            AmendmentFileNames = new List<string>();
        }
    }
}
