using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;

namespace docx2txt
{
    class DocxConverter
    {
        private Application _word { get; } = new Application();
            
        public void Convert(string input, string output)
        {
            var doc = _word.Documents.Open(input);
            
            doc.SaveAs2(
                output, 
                FileFormat: WdSaveFormat.wdFormatEncodedText, 
                Encoding: Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8);

            doc.Close();
        }
    }
}
