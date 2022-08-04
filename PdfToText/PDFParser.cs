using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace PdfToText
{
    /// <summary>
    /// Parses a PDF file and extracts the text from it.
    /// </summary>
    public class PDFParser
    {

        /// BT = Beginning of a text object operator 
        /// ET = End of a text object operator
        /// Td move to the start of next line
        ///  5 Ts = superscript
        /// -5 Ts = subscript

        #region Fields

        #region _numberOfCharsToKeep
        /// <summary>
        /// The number of characters to keep, when extracting text.
        /// </summary>
        private static int _numberOfCharsToKeep = 15;
        #endregion

        #endregion

        #region ExtractText  


        public void ExtractText(string inFileName)
        {
            using (Stream newpdfStream = new FileStream(inFileName, FileMode.Open, FileAccess.ReadWrite))
            {
                PdfReader pdfReader = new PdfReader(newpdfStream);
                ///string fname = "C:\\c#\\PDF Util\\jul2019_asuri.txt";
                string fname="C:\\payslip\\process\\july2022.txt";
                StringBuilder errpages = new StringBuilder();
                //int pageSize = (int)Math.Ceiling((double)pdfReader.NumberOfPages / (double)(Environment.ProcessorCount * 2));
                //int numberOfThreads = (int)Math.Ceiling((double)pdfReader.NumberOfPages / (double)pageSize);
                IList<Task> tasks = new List<Task>();
                //for (int index = 0; index < numberOfThreads; index++)
                for (int index = 1; index <= pdfReader.NumberOfPages; index++)
                {
                    //int currentIndex = index;
                    var pageno = index;
                    var subIndex = index;
                    //int page = Math.Min((index + 1) * pageSize, pdfReader.NumberOfPages);
                    tasks.Add(Task.Factory.StartNew(() =>
                    {
                        StringBuilder taskResult = new StringBuilder();
                        //for (int subIndex = (currentIndex * pageSize + 1); subIndex <= page; subIndex++)
                        //{
                        taskResult.Clear();
                        taskResult.AppendLine(PdfTextExtractor.GetTextFromPage(pdfReader, subIndex, new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy()));
                        //var txt = taskResult.ToString();
                        string[] lines = taskResult.ToString().Split(Environment.NewLine.ToCharArray());
                        
                        if(lines.Length>30)
                        {

                            //var date = lines[4].ToString().Replace("-", "");
                            //var temp = new String(lines[6].Where(Char.IsDigit).ToArray());
                            //var pathstr = string.Format($"c:\\payslip\\{date}\\");
                            //if (!Directory.Exists(pathstr))
                            //   Directory.CreateDirectory(pathstr);
                            //getUnionMembers(lines, fname, pageno);
                            //lock{
                            //getStaffDetails(lines, fname);
                           //getArrears(lines, fname, pageno);
                            GetPageNumber(lines, fname, pageno);
                            //}

                        }
                        else
                        {
                            errpages.AppendLine(string.Format($"Error on page: {subIndex.ToString()}"));
                            Console.WriteLine(string.Format($"Error on page: {subIndex.ToString()}"));
                        }
                        
                        //var outputPdfPath = string.Format($"{pathstr}{temp}_{date}_payslip.pdf");
                        ///CreatePDF(pdfReader, subIndex, outputPdfPath);

                        //pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

                        //}
                        //return taskResult.ToString();
                    }));
                       
                    //.ContinueWith((t) => File.WriteAllText(fname + ".txt", t.Result)));
                }

                Task.WaitAll(tasks.ToArray());
                Console.WriteLine("Finish");
                using (StreamWriter fs = new StreamWriter(fname, true))
                {
                    fs.Write(usb.ToString());
                }

                using (StreamWriter fs = new StreamWriter("C:\\payslip\\errorpages.txt", true))
                {
                    fs.Write(errpages.ToString());
                }
            }
        }

        StringBuilder usb = new StringBuilder();
        private void GetPageNumber(string[] values, string fname, int pageno)
        {
            string ippis = "", grade = "", name = "";
            string netpay = "";
            string step = "";
            foreach (var v in values)
            {
               
                if (v.ToString().Contains("IPPIS  Number:"))
                {
                    ippis = new String(v.Where(Char.IsDigit).ToArray()).Trim();

                }
                if (v.ToString().Contains("Employee Name:"))
                {
                    name = new String(v.ToArray());
                }
                 if (v.Contains("Grade:"))
                 {
                     grade = new String(v.Where(char.IsDigit).ToArray()).Trim();
                 }
                if (v.Contains("Step:"))
                {
                    step = new String(v.Where(char.IsDigit).ToArray()).Trim();
                }
                if (v.Contains("Total Gross Earnings :"))
                {
                    netpay = new String(v.Where(char.IsDigit).ToArray()).Trim();
                }
            }
            //if (Convert.ToInt32(grade) > 13)
            //{
            //    var final = $"{ippis}|{netpay}";
            //    usb.AppendLine(final);
            //}
            var final = $"{ippis}|{pageno}|{name}";
            //var final = $"{ippis}|{grade}|{step}|{netpay}";
            //var final = $"{ippis}|{pageno}|{grade}|{step}|{netpay}";
            usb.AppendLine(final);
        }
        private void getArrears(string[] values, string fname, int pageno)
        {
            string ippis = "", grade = "", arrears="";
            foreach(var v in values)
            {
                if (v.ToString().Contains("MinimumWage Arrears"))
                {
                    arrears = new String(v.Where(Char.IsDigit).ToArray()).Trim();
                }
                if (v.ToString().Contains("IPPIS  Number:"))
                {
                    ippis = new String(v.Where(Char.IsDigit).ToArray()).Trim();

                }
                if (v.Contains("Grade:"))
                {
                    grade = new String(v.Where(char.IsDigit).ToArray()).Trim();
                }
            }
            if (!string.IsNullOrEmpty(arrears))
            {
                var minimumWageArrears = string.Format($"{ippis}|{pageno}");
                usb.AppendLine(minimumWageArrears);
            }
        }
        private void getUnionMembers(string[] values, string fname, int pageno)
        {
            string union = "", ippis = "", dues = "", grade="";
            foreach (var v in values)
            {
                if (v.ToString().Contains("IPPIS  Number:"))
                {
                    ippis = new String(v.Where(Char.IsDigit).ToArray()).Trim();

                }
                if (v.Contains("Trade Union: ACADEMIC STAFF UNION OF RESEARCH INST"))
                {
                    union = v.Trim();
                }
                if (v.Contains("CONRAISS UNION DUES"))
                {
                    dues = v.Replace("CONRAISS UNION DUES","").Trim();// new String(v.Where(Char.IsDigit).ToArray());
                }
                /*if (v.Contains("Consequential Union"))
                {
                    dues = v.Replace("Consequential Union", "").Trim();
                }*/
                if (v.Contains("Grade:"))
                {
                    grade = new String(v.Where(char.IsDigit).ToArray()).Trim();
                }

            }
            if(!string.IsNullOrEmpty(union))
            {
                //var final = string.Format($"{ippis}|{dues}|{grade}");
                var final = string.Format($"{ippis}|{pageno}");
                usb.AppendLine(final);
            }
            
            //var text = usb.ToString();
            
            //File.WriteAllText(fname + ".txt", text);
        }
        private void getStaffDetails(string[] values, string fname)
        {
            string ippis = "", name = "", tax = "", grade="", step="", grosspay="", account_no="", bank_name="";
            foreach (var v in values)
            {
                if (v.ToString().Contains("IPPIS  Number:"))
                {
                    ippis = new String(v.Where(Char.IsDigit).ToArray()).Trim();

                }
                if (v.ToString().Contains("Employee Name:"))
                {
                    name = new String(v.ToArray());
                }
                /*if (v.ToString().Contains("Tax State"))
                {
                    tax = new String(v.ToArray());
                }*/
                if (v.ToString().Contains("Grade:"))
                {
                    grade = new String(v.Where(Char.IsDigit).ToArray()).Trim();
                }
                if (v.ToString().Contains("Step:"))
                {
                    step = new String(v.ToArray());
                }
                if(v.ToString().Contains("CONRAISS Cons Salary02"))
                {
                    grosspay = new String(v.ToArray());
                }
                if (v.ToString().Contains("Bank Name:"))
                {
                    bank_name = new String(v.ToArray());
                }
                if (v.ToString().Contains("Account Number:"))
                {
                    account_no = new String(v.Where(Char.IsDigit).ToArray()).Trim();
                }
                

            }

            
            StringBuilder sb = new StringBuilder();
            /* sb.Append(string.Format($" \"{ippis}\"," ));
             sb.Append(string.Format($" \"{grade}\","));

             usb.AppendLine(sb.ToString());*/

            var final = string.Format($"{ippis}|{grade}|{step}|{grosspay}|{account_no}|{bank_name})|{name}");
            //var final = string.Format($"{ippis}|{ grade}|{ step}|{ name}");
            usb.AppendLine(final);


            //var text = usb.ToString();

            //File.WriteAllText(fname + ".txt", text);
        }

        private void CreatePDF(PdfReader pdfReader, int page, string outputPdfPath)
        {
            using (Document sourceDocument = new Document(pdfReader.GetPageSizeWithRotation(page)))
            {
                using (PdfCopy pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create)))
                {
                    sourceDocument.Open();
                    PdfImportedPage importedPage = pdfCopyProvider.GetImportedPage(pdfReader, page);
                    pdfCopyProvider.AddPage(importedPage);
                }
            }
        }

        /// <summary>
        /// Extracts a text from a PDF file.
        /// </summary>
        /// <param name="inFileName">the full path to the pdf file.</param>
        /// <param name="outFileName">the output file name.</param>
        /// <returns>the extracted text</returns>
        public bool ExtractText(string inFileName, string outFileName)
        {
            StringBuilder sb = new StringBuilder();
            StreamWriter outFile = null;
            try
            {
                if (!File.Exists(inFileName))
                    return false;
                // Create a reader for the given PDF file
                PdfReader reader = new PdfReader(inFileName);

                //var sText = PdfTextExtractor.GetTextFromPage(reader, 1);


                //outFile = File.CreateText(outFileName);
                outFile = new StreamWriter(outFileName, false, System.Text.Encoding.UTF8);

                Console.Write("Processing: ");

                int totalLen = 68;
                float charUnit = ((float)totalLen) / (float)reader.NumberOfPages;
                int totalWritten = 0;
                float curUnit = 0;

                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                //reader.RemoveUnusedObjects;
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {

                    string currentText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);

                    currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    sb.Append(currentText);
                    var t = sb.ToString();


                    string thePage = PdfTextExtractor.GetTextFromPage(reader, page, strategy);
                    string[] theLines = thePage.Split('\n');
                    foreach (var theLine in theLines)
                    {
                        sb.AppendLine(theLine);
                    }
                    var kk = sb.ToString();
                    /*string txt = PdfTextExtractor.GetTextFromPage(reader, page + 1, strategy);
                    if (!string.IsNullOrWhiteSpace(txt))
                    {
                        sb.Append(Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(txt))));
                    }*/



                    var text = ExtractTextFromPDFBytes(reader.GetPageContent(page));
                    outFile.Write(text + " ");

                    // Write the progress.
                    if (charUnit >= 1.0f)
                    {
                        for (int i = 0; i < (int)charUnit; i++)
                        {
                            Console.Write("#");
                            totalWritten++;
                        }
                    }
                    else
                    {
                        curUnit += charUnit;
                        if (curUnit >= 1.0f)
                        {
                            for (int i = 0; i < (int)curUnit; i++)
                            {
                                Console.Write("#");
                                totalWritten++;
                            }
                            curUnit = 0;
                        }

                    }
                }

                if (totalWritten < totalLen)
                {
                    for (int i = 0; i < (totalLen - totalWritten); i++)
                    {
                        Console.Write("#");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                if (outFile != null) outFile.Close();
            }
        }
        #endregion

        #region ExtractTextFromPDFBytes
        /// <summary>
        /// This method processes an uncompressed Adobe (text) object 
        /// and extracts text.
        /// </summary>
        /// <param name="input">uncompressed</param>
        /// <returns></returns>
        private string ExtractTextFromPDFBytes(byte[] input)
        {
            if (input == null || input.Length == 0) return "";

            try
            {
                string resultString = "";

                // Flag showing if we are we currently inside a text object
                bool inTextObject = false;

                // Flag showing if the next character is literal 
                // e.g. '\\' to get a '\' character or '\(' to get '('
                bool nextLiteral = false;

                // () Bracket nesting level. Text appears inside ()
                int bracketDepth = 0;

                // Keep previous chars to get extract numbers etc.:
                char[] previousCharacters = new char[_numberOfCharsToKeep];
                for (int j = 0; j < _numberOfCharsToKeep; j++) previousCharacters[j] = ' ';


                for (int i = 0; i < input.Length; i++)
                {
                    char c = (char)input[i];

                    if (inTextObject)
                    {
                        // Position the text
                        if (bracketDepth == 0)
                        {
                            if (CheckToken(new string[] { "TD", "Td" }, previousCharacters))
                            {
                                resultString += "\n\r";
                            }
                            else
                            {
                                if (CheckToken(new string[] { "'", "T*", "\"" }, previousCharacters))
                                {
                                    resultString += "\n";
                                }
                                else
                                {
                                    if (CheckToken(new string[] { "Tj" }, previousCharacters))
                                    {
                                        resultString += " ";
                                    }
                                }
                            }
                        }

                        // End of a text object, also go to a new line.
                        if (bracketDepth == 0 &&
                            CheckToken(new string[] { "ET" }, previousCharacters))
                        {

                            inTextObject = false;
                            resultString += " ";
                        }
                        else
                        {
                            // Start outputting text
                            if ((c == '(') && (bracketDepth == 0) && (!nextLiteral))
                            {
                                bracketDepth = 1;
                            }
                            else
                            {
                                // Stop outputting text
                                if ((c == ')') && (bracketDepth == 1) && (!nextLiteral))
                                {
                                    bracketDepth = 0;
                                }
                                else
                                {
                                    // Just a normal text character:
                                    if (bracketDepth == 1)
                                    {
                                        // Only print out next character no matter what. 
                                        // Do not interpret.
                                        if (c == '\\' && !nextLiteral)
                                        {
                                            nextLiteral = true;
                                        }
                                        else
                                        {
                                            if (((c >= ' ') && (c <= '~')) ||
                                                ((c >= 128) && (c < 255)))
                                            {
                                                resultString += c.ToString();
                                            }

                                            nextLiteral = false;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Store the recent characters for 
                    // when we have to go back for a checking
                    for (int j = 0; j < _numberOfCharsToKeep - 1; j++)
                    {
                        previousCharacters[j] = previousCharacters[j + 1];
                    }
                    previousCharacters[_numberOfCharsToKeep - 1] = c;

                    // Start of a text object
                    if (!inTextObject && CheckToken(new string[] { "BT" }, previousCharacters))
                    {
                        inTextObject = true;
                    }
                }
                return resultString;
            }
            catch
            {
                return "";
            }
        }
        #endregion

        #region CheckToken
        /// <summary>
        /// Check if a certain 2 character token just came along (e.g. BT)
        /// </summary>
        /// <param name="search">the searched token</param>
        /// <param name="recent">the recent character array</param>
        /// <returns></returns>
        private bool CheckToken(string[] tokens, char[] recent)
        {
            foreach (string token in tokens)
            {
                if ((recent[_numberOfCharsToKeep - 3] == token[0]) &&
                    (recent[_numberOfCharsToKeep - 2] == token[1]) &&
                    ((recent[_numberOfCharsToKeep - 1] == ' ') ||
                    (recent[_numberOfCharsToKeep - 1] == 0x0d) ||
                    (recent[_numberOfCharsToKeep - 1] == 0x0a)) &&
                    ((recent[_numberOfCharsToKeep - 4] == ' ') ||
                    (recent[_numberOfCharsToKeep - 4] == 0x0d) ||
                    (recent[_numberOfCharsToKeep - 4] == 0x0a))
                    )
                {
                    return true;
                }
            }
            return false;
        }
        #endregion
    }
}
