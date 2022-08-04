using System;
using System.Text;
using System.IO;

namespace PdfToText
{
    /// <summary>
    /// The main entry point to the program.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                /* if (args.Length < 1)
                 {
                     DisplayUsage();
                     return;
                 }*/

                string file = "c:\\payslip\\2022\\july2022.pdf";
                //string file = "c:\\works\\2020\\payslip\\apr2020\\apr2020.pdf";
                //string file = @"C:\Payslip\process\april2022\ict.pdf";

                if (File.Exists(file))
                {
                    file = Path.GetFullPath(file);
                    if (!File.Exists(file))
                    {
                        Console.WriteLine("Please give in the path to the PDF file.");
                    }
                }

                PDFParser pdfParser = new PDFParser();
                pdfParser.ExtractText(file);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc);
            }
        }

        static void DisplayUsage()
        {
            Console.WriteLine();
            Console.WriteLine("Usage:\tpdftotext FILE");
            Console.WriteLine();
            Console.WriteLine("\tFILE\t the path to the PDF file, it may be relative or absolute.");
            Console.WriteLine();
        }
    }
}
