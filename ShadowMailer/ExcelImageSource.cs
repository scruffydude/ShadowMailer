using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShadowMailer
{
    public class ExcelImageSource
    {
        private static string _filePath;
        private static string[] _imageRange;

        public string FilePath
        {
            get { return _filePath; }
            set { _filePath = value; }
        }
        public string[] ImageRanges
        {
            get { return _imageRange; }
            set { _imageRange = value; }
        }
    }
}
