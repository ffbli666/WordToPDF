using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;

namespace WordToPDF
{
    class Program
    {
        static int Main(string[] args)
        {
            int result = 0;
            string input = "";
            string output = "";

            if (args.Length == 0)
            {
                Console.WriteLine("Version: WordToPDF v0.1");
                Console.WriteLine("Author: ZongYing Lyu");
                Console.WriteLine("Usage: ConvertPDF <input file> <output file>");
                return 1;
            }

            try
            {
                input = Path.GetFullPath(args[0]);
                output = Path.GetDirectoryName(input) + "\\" + Path.GetFileNameWithoutExtension(input) + ".pdf";
                if (args.Length >= 2)
                {
                    output = Path.GetFullPath(args[1]);
                    if (Path.GetFileName(output) == "" || !Path.HasExtension(output))
                    {
                        output = Path.GetDirectoryName(output) + "\\" + Path.GetFileNameWithoutExtension(input) + ".pdf";
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Input or Ouput has error path");
                Console.WriteLine(e);
                return 1;

            }
            //Console.WriteLine("input:{0}", input);
            //Console.WriteLine("output:{0}", output);
            if (!File.Exists(input))
            {
                Console.WriteLine("Input file does not exist: {0}", input);
                return 1;
            }

            string ext = Path.GetExtension(input).ToLower();
            switch (ext)
            {
                case ".doc":
                case ".docx":
                    result = WordToPDF(input, output);
                    break;
                case ".ppt":
                case ".pptx":
                    result = PPToPDF(input, output);
                    break;
                default:
                    Console.WriteLine("Only support .doc, .docx, .ppt, .pptx", input, output);
                    return 1;
            }
            if (result == 0)
            {
                Console.WriteLine("Convert success: {0}", output);
            }
            return result;
        }

        private static int WordToPDF(string input = "", string output = "")
        {

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            try
            {
                Microsoft.Office.Interop.Word.Document wordDocument = appWord.Documents.Open(input);
                wordDocument.ExportAsFixedFormat(output, WdExportFormat.wdExportFormatPDF);
                wordDocument.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Can not open input file: {0}", input);
                Console.WriteLine(e);
                return 1;
            }
            return 0;
        }

        private static int PPToPDF(string input = "", string output = "")
        {

            Microsoft.Office.Interop.PowerPoint.Application appPPT = new Microsoft.Office.Interop.PowerPoint.Application();
            try
            {
                Microsoft.Office.Interop.PowerPoint.Presentation pptPst = appPPT.Presentations.Open(input,
                                                                                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                                                    Microsoft.Office.Core.MsoTriState.msoFalse
                                                                                                    );
                pptPst.ExportAsFixedFormat(output, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                pptPst.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Can not open input file: {0}", input);
                Console.WriteLine(e);
                return 1;
            }
            return 0;
        }
    }
}
