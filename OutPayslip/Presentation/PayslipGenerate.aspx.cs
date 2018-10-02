using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using OutPayslip.DataTransferObject;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace OutPayslip.Presentation
{
    public partial class PayslipGenerate : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        public void Generate(object sender, EventArgs e)
        {
            if (Fileupload.HasFile)
            {
                string FileName = Path.GetFileNameWithoutExtension(Fileupload.FileName);
                string FileExt = Path.GetExtension(Fileupload.FileName);
                string Filepath = CreateFile(FileName, FileExt, null);
                if (FileExt.ToLower() == ".xlsx")
                {
                    try
                    {

                        DataSet ExcelData = new DataSet();
                        ExcelData = Services.ExcelReader.CreateDataset(Filepath);
                        List<Document> documents = CovertDataSetToPaySlip(ExcelData);

                    }
                    catch (Exception E)
                    {
                        throw E;
                    }
                }
            }

        }
        public string AppendComma(string data)
        {
            if (data == "")
            {
                return "";
            }
            else
            {
                return String.Format("{0:#,0.00}", Convert.ToDecimal(data));
            }
        }
        public string GenerateWordsinRs(string inputRs)
        {
            string input = inputRs;
            string a = "";
            string b = "";

            // take decimal part of input. convert it to word. add it at the end of method. 
            string decimals = "";

            if (input.Contains("."))
            {
                decimals = input.Substring(input.IndexOf(".") + 1);
                // remove decimal part from input 
                input = input.Remove(input.IndexOf("."));

            }
            string strWords = NumbersToWords(Convert.ToInt32(input));

            if (!inputRs.Contains("."))
            {
                a = strWords + " Rupees Only";
            }
            else
            {
                a = strWords + " Rupees";
            }

            if (decimals.Length > 0)
            {
                // if there is any decimal part convert it to words and add it to strWords. 
                string strwords2 = NumbersToWords(Convert.ToInt32(decimals));
                b = " and " + strwords2 + " Paisa Only ";
            }

            string final2 = "";
            final2 = a + b;
            return final2;
        }

        public static string NumbersToWords(int inputNumber)
        {
            int inputNo = inputNumber;

            if (inputNo == 0)
                return "Zero";

            int[] numbers = new int[4];
            int first = 0;
            int u, h, t;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            if (inputNo < 0)
            {
                sb.Append("Minus ");
                inputNo = -inputNo;
            }

            string[] words0 = {"" ,"One ", "Two ", "Three ", "Four ",
 "Five " ,"Six ", "Seven ", "Eight ", "Nine "};
            string[] words1 = {"Ten ", "Eleven " , "Twelve ", "Thirteen " , "Fourteen ",
 "Fifteen ", "Sixteen " ,"Seventeen ", "Eighteen " , "Nineteen "};
            string[] words2 = {"Twenty ", "Thirty " , "Forty ", "Fifty ", "Sixty ",
 "Seventy ", "Eighty " , "Ninety "};
            string[] words3 = { "Thousand ", "Lakh ", "Crore " };

            numbers[0] = inputNo % 1000; // units 
            numbers[1] = inputNo / 1000;
            numbers[2] = inputNo / 100000;
            numbers[1] = numbers[1] - 100 * numbers[2]; // thousands 
            numbers[3] = inputNo / 10000000; // crores 
            numbers[2] = numbers[2] - 100 * numbers[3]; // lakhs 

            for (int i = 3; i > 0; i--)
            {
                if (numbers[i] != 0)
                {
                    first = i;
                    break;
                }
            }
            for (int i = first; i >= 0; i--)
            {
                if (numbers[i] == 0) continue;
                u = numbers[i] % 10; // ones 
                t = numbers[i] / 10;
                h = numbers[i] / 100; // hundreds 
                t = t - 10 * h; // tens 
                if (h > 0) sb.Append(words0[h] + "Hundred and ");
                if (u > 0 || t > 0)
                {
                    if (h > 0 || i == 0) sb.Append("");
                    if (t == 0)
                        sb.Append(words0[u]);
                    else if (t == 1)
                        sb.Append(words1[u]);
                    else
                        sb.Append(words2[t - 2] + words0[u]);
                }
                if (i != 0) sb.Append(words3[i - 1]);
            }
            return sb.ToString().TrimEnd();
        }
        private string CreateFile(string fileName, string extension, Document Doc)
        {

            string processFilePath = string.Empty,
                pathToSaveDecryptedFile = string.Empty;
            try
            {

                string saveFile = string.Format("{0}{1}{2}{3}",
                              Server.MapPath("FileToSave").ToString(),
                              fileName, DateTime.Now.ToString("MMddyyyy_hhmmss"), extension);
                if (extension == ".xlsx")
                {
                    Fileupload.SaveAs(saveFile);
                }
                else if (extension == ".pdf")
                {
                    PdfWriter.GetInstance(Doc, new FileStream(saveFile, FileMode.Create));
                }
                pathToSaveDecryptedFile = string.Format("{0}{1}{2}{3}{4}{5}",
                    Path.GetDirectoryName(saveFile), "//", Path.GetFileNameWithoutExtension(saveFile), "_"
                    , "Decrypt", extension);
                processFilePath = saveFile;

            }
            catch (Exception E)
            {
                throw E;
            }
            return processFilePath;

        }
        public string GetMonth(int MonthInt)
        {
            string ExactMonth = string.Empty;
            string[] Months = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            if (MonthInt <= 12)
            {
                ExactMonth = Months[MonthInt - 1];
            }
            return ExactMonth;
        }
        public List<Document> CovertDataSetToPaySlip(DataSet ds)
        {
            try
            {
                int CurrentMonth = DateTime.Now.Month;
                int CurrentYear = DateTime.Now.Year;
                string Month = GetMonth(CurrentMonth-1);
                List<Document> documents = new List<Document>();
                for (int j = 0; j < ds.Tables.Count; j++)
                {
                    if (ds != null && ds.Tables[j].Columns.Count > 0 && ds.Tables[j].Rows.Count > 0)
                    {
                        for (int k = 0; k <= ds.Tables[j].Rows.Count - 1; k++)
                        {
                            Document doc = new Document(PageSize.A4, 36f, 36f, 36f, 36f);//36f, 36f, 90f, 100f); 
                            using (MemoryStream memorystream = new MemoryStream())
                            {
                                PdfWriter.GetInstance(doc, memorystream);
                                doc.Open();

                                // 1) Adding logo to right side top 
                                PdfPTableHeader pTableHeader = new PdfPTableHeader();
                                string imagePath = Server.MapPath("images") + "\\logo.png";
                                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(imagePath);
                                image.Alignment = Element.ALIGN_LEFT;
                                // set width and height 
                                image.ScaleToFit(180f, 250f);
                                Paragraph title = new Paragraph
                                {
                            new Chunk(@"MOXIETEK E&I SERVICES", new Font(Font.FontFamily.TIMES_ROMAN, 11, 1, BaseColor.BLACK)),
                                };
                                Paragraph title1 = new Paragraph
                                {
                            new Chunk(@"9 / 284, First Ambal Nagar, Thiruvalluvar street, Kovur, Chennai, 600122", new Font(Font.FontFamily.TIMES_ROMAN, 8, 1, BaseColor.BLACK)),
                                };

                                // 4) Addling blank paragraph 
                                doc.Add(new Paragraph("  "));

                                Paragraph title2 = new Paragraph
                                {
                            new Chunk("Pay Slip for the month of " + Month +" "+CurrentYear, new Font(Font.FontFamily.TIMES_ROMAN, 10, Font.BOLD, BaseColor.BLACK))
                                };
                                title.Alignment = 1;
                                title.Alignment = Element.ALIGN_CENTER;
                                title1.Alignment = 1;
                                title2.Alignment = 1;
                                title1.Alignment = Element.ALIGN_CENTER;
                                title2.Alignment = Element.ALIGN_CENTER;
                                doc.Add(image);
                                doc.Add(title);
                                doc.Add(title1);
                                doc.Add(title2);

                                // 4) Addling blank paragraph 
                                doc.Add(new Paragraph("  "));

                                // 5) Creating 1st table with 4 column 
                                PdfPTable table1 = new PdfPTable(4);
                                table1.DefaultCell.Border = Rectangle.NO_BORDER;
                                int[] columnwidth = { 20, 25, 20, 25 };
                                table1.SetWidths(columnwidth);
                                table1.WidthPercentage = 100;
                                table1.HorizontalAlignment = 0;

                                // 6) Adding employee data to table1 
                                for (int i = 0; i <= 13; i++)
                                {
                                    string columnName = (ds.Tables[j].Columns[i].ColumnName);
                                    string columnValue = (ds.Tables[j].Rows[k][i].ToString());

                                    PdfPCell cellColumnName = new PdfPCell(new Phrase(columnName, new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) };
                                    cellColumnName.HorizontalAlignment = 0;
                                    table1.AddCell(cellColumnName);

                                    PdfPCell cellColumnValue = new PdfPCell(new Phrase(columnValue == "" ? "" : columnValue, new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK)));
                                    cellColumnValue.HorizontalAlignment = 0;//0=Left, 1=Centre, 2=Right 
                                    table1.AddCell(cellColumnValue);
                                }
                                doc.Add(table1);

                                // 4) Addling blank paragraph 
                                doc.Add(new Paragraph("  "));

                                PdfPTable mainTable = new PdfPTable(4); //earnedTable1.TotalWidth = 500f;//earnedTable1.LockedWidth = true; 
                                int[] columnwidth1 = { 30, 20, 30, 20 }; //23, 20, 25, 32 }; 
                                mainTable.SetWidths(columnwidth1);
                                mainTable.WidthPercentage = 100;
                                mainTable.HorizontalAlignment = 0;

                                // a. adding 4 cells for header 
                                mainTable.AddCell(new PdfPCell(new Phrase("Earnings", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) }); //{ HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, BackgroundColor = new BaseColor(System.Drawing.Color.Silver) };; 
                                mainTable.AddCell(new PdfPCell(new Phrase("Amount(Rs)", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) }); //{ HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, BackgroundColor = new BaseColor(System.Drawing.Color.Silver) };; 
                                mainTable.AddCell(new PdfPCell(new Phrase("Deduction", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) }); //{ HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, BackgroundColor = new BaseColor(System.Drawing.Color.Silver) };; 
                                mainTable.AddCell(new PdfPCell(new Phrase("Amount(Rs)", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) }); //{ HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, BackgroundColor = new BaseColor(System.Drawing.Color.Silver) };; 

                                // b. creating earning table with 2 columns [left side] 
                                PdfPTable earning = new PdfPTable(2);
                                int[] columnwidth3 = { 30, 20 };
                                earning.SetWidths(columnwidth3);
                                earning.WidthPercentage = 80;
                                earning.HorizontalAlignment = 0;

                                // c. adding earning data 
                                for (int i = 14; i <= 16; i++)
                                {
                                    string columnName = (ds.Tables[j].Columns[i].ColumnName);
                                    string columnValue = (ds.Tables[j].Rows[k][i].ToString());

                                    PdfPCell cellColumnName = new PdfPCell(new Phrase(columnName, new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK)));
                                    cellColumnName.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right 
                                    earning.AddCell(cellColumnName);

                                    PdfPCell cellColumnValue = new PdfPCell(new Phrase(columnValue == "" ? "" : AppendComma(columnValue), new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT };
                                    earning.AddCell(cellColumnValue);
                                }

                                // d. creating deduction table with 2 columns [Right side] 
                                PdfPTable deduction = new PdfPTable(2);
                                int[] columnwidth4 = { 30, 20 };
                                deduction.SetWidths(columnwidth3);
                                deduction.WidthPercentage = 80;
                                deduction.HorizontalAlignment = 0;

                                // e. adding deduction data 
                                for (int i = 18; i <= 22; i++)
                                {
                                    string columnName = (ds.Tables[j].Columns[i].ColumnName);
                                    string columnValue = (ds.Tables[j].Rows[k][i].ToString());

                                    PdfPCell cellColumnName = new PdfPCell(new Phrase(columnName, new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK)));
                                    cellColumnName.HorizontalAlignment = 0;
                                    cellColumnName.Colspan = 1;//0=Left, 1=Centre, 2=Right 
                                    deduction.AddCell(cellColumnName);

                                    PdfPCell cellColumnValue = new PdfPCell(new Phrase(columnValue == "" ? "" : AppendComma(columnValue), new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT };
                                    cellColumnValue.Colspan = 1;
                                    deduction.AddCell(cellColumnValue);
                                }
                                // f. creating a new cell [cell1] with colspan=2 
                                //    adding earning table into cell1 
                                PdfPCell cell1 = new PdfPCell(earning);
                                cell1.Colspan = 2;
                                // adding cell1 into mainTable 
                                mainTable.AddCell(cell1);

                                // g. creating a new cell [cell2] with colspan=2 
                                //    adding deduction table into cell2 
                                PdfPCell cell2 = new PdfPCell(deduction);
                                cell2.Colspan = 2;
                                // adding cell2 into mainTable 
                                mainTable.AddCell(cell2);

                                mainTable.AddCell(new PdfPCell(new Phrase("Total Earning", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, BackgroundColor = new BaseColor(236, 236, 236) });
                                mainTable.AddCell(new PdfPCell(new Phrase(AppendComma(ds.Tables[j].Rows[k][17].ToString()), new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5 });
                                mainTable.AddCell(new PdfPCell(new Phrase("Total Deductions", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5, BackgroundColor = new BaseColor(236, 236, 236) });
                                mainTable.AddCell(new PdfPCell(new Phrase(AppendComma(ds.Tables[j].Rows[k][23].ToString()), new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5 });

                                PdfPCell netEarning = new PdfPCell(new Phrase("Net Salary", new Font(Font.FontFamily.TIMES_ROMAN, 10, Font.BOLD, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, BackgroundColor = new BaseColor(236, 236, 236) };
                                mainTable.AddCell(netEarning);

                                string NetSalary = (Convert.ToDecimal(ds.Tables[j].Rows[k][24].ToString()).ToString());
                                PdfPCell netSalary = new PdfPCell(new Phrase(AppendComma(NetSalary), new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_RIGHT, Padding = 5 };
                                mainTable.AddCell(netSalary);

                                PdfPCell blankCell = new PdfPCell();
                                blankCell.Colspan = 2;
                                mainTable.AddCell(blankCell);

                                PdfPCell NetSalaryInWords = new PdfPCell(new Phrase("Net Salary In Word : " + GenerateWordsinRs(NetSalary), new Font(Font.FontFamily.TIMES_ROMAN, 8, Font.BOLD, BaseColor.BLACK))) { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5 };
                                NetSalaryInWords.Colspan = 4;
                                mainTable.AddCell(NetSalaryInWords);

                                // adding mainTable to document object 
                                doc.Add(mainTable);

                                Paragraph Note = new Paragraph();
                                Note.Add(new Chunk("This is computer generated payslip and does not require signature or company seal.", new Font(Font.FontFamily.TIMES_ROMAN, 10, 1, BaseColor.BLACK)));
                                Note.Alignment = 1;
                                doc.Add(Note);

                                doc.Close();
                                byte[] bytes = memorystream.ToArray();
                                memorystream.Close();
                                string fileName = " " + Month + ".pdf";
                                string Email = (ds.Tables[j].Rows[k][25]).ToString();
                                MailMessage mail = new MailMessage();
                                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                                mail.From = new MailAddress("r.deepakram06@gmail.com");
                                mail.To.Add(Email);
                                mail.Subject = "Payslip For " + Month;
                                mail.Body = "Hi,";
                                mail.Body = "";
                                mail.Body = "This is computer generated payslip attached in the mail, " +
                                                 "Please Refer the Attachement for further clarification feel free to contact";
                                mail.IsBodyHtml = false;
                                System.Net.Mail.Attachment attachment;
                                attachment = new System.Net.Mail.Attachment(new MemoryStream(bytes), fileName);
                                mail.Attachments.Add(attachment);
                                SmtpServer.UseDefaultCredentials = false;
                                SmtpServer.Credentials = new System.Net.NetworkCredential("r.deepakram06@gmail.com", "sathyaDeepak");

                                SmtpServer.EnableSsl = true;
                                SmtpServer.Port = 587;
                                SmtpServer.Send(mail);
                                // documents.Add(doc);
                            }
                        }
                    }

                }

                return documents;
            }
            catch (Exception E)
            {
                throw E;
            }
        }
    }
}