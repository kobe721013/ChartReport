using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using MiscUtils;

namespace ConsoleApp4GenerateChart
{
    class MyReportOutputter
    {
        private Dictionary<string, double[]> tableDataSets = new Dictionary<string, double[]>();
        private Dictionary<string, double[]> chartDataSets = new Dictionary<string, double[]>();


        private string _DefaultAccessPath = Directory.GetCurrentDirectory();
        private string _TableFileNameKey = @"OpsCenter_Deduplication_Rates_by_Policy_Type_";
        private string _ChartFileNameKey = @"OpsCenter_last_month_disk_usage_";
        private Color _ChartSerieColor = Color.FromArgb(153, 217, 234);
        private BaseFont bfChinese = BaseFont.CreateFont("C:\\Windows\\Fonts\\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        public void StartBy(string date, string searchDocPath, string outputPath)
        {
            try
            {
                startParse(date, searchDocPath, outputPath);
            }
            catch (Exception e)
            {
                Log.Error($"StartBy fail. ErrorMessage={e.Message}. Stacktrace={e.StackTrace}");
            }
            finally
            {
                cleanTemoFolder();
            }
        }

        public void StartBy(string date)
        {
            try
            {
                startParse(date, _DefaultAccessPath, _DefaultAccessPath);
            }
            catch (Exception e)
            {

                Log.Error($"StartBy fail. ErrorMessage={e.Message}. Stacktrace={e.StackTrace}");
            }
            finally
            {
                cleanTemoFolder();
            }
        }

        private void cleanTemoFolder()
        {
            DirectoryInfo dir = new DirectoryInfo(_DefaultAccessPath);


            string pattern = $"*_imgtmp";
            DirectoryInfo[] dirs = dir.GetDirectories(pattern, SearchOption.TopDirectoryOnly);

            foreach (DirectoryInfo d in dirs)
            {
                d.Delete(true);
            }

        }

        private bool startParse(string date, string docPath, string outputPath)
        {
            DateTime dateTime = DateTime.Parse(date);

            parseTableFiles(docPath, dateTime);
            parseChartFiles(docPath, dateTime);



            //prepare char and temp folder
            string tempFolder = DateTime.Now.ToString("yyyyMMddhhmmss") + "_imgtmp";
            string chartTempPath = Path.Combine(Directory.GetCurrentDirectory(), tempFolder);
            Directory.CreateDirectory(chartTempPath);
            foreach (KeyValuePair<string, double[]> item in chartDataSets)
            {
                createChart(item.Key, item.Value, Path.Combine(chartTempPath, $"{item.Key}.png"));
            }

            //prepare pdfTable
            Dictionary<string, PdfPTable> pdfTables = new Dictionary<string, PdfPTable>();
            foreach (KeyValuePair<string, double[]> item in tableDataSets)
            {
                var tableData = createTable(item.Key, item.Value);
                pdfTables.Add(item.Key, tableData);
            }

            save2Pdf(pdfTables, chartTempPath, outputPath);

            return true;
        }

        private void save2Pdf(Dictionary<string, PdfPTable> pdfTables, string chartTempPath, string outputPath)
        {
            Document doc = new Document(PageSize.A4, 50, 50, 80, 50);
            try
            {

                FileStream fs = new FileStream(outputPath, FileMode.Create);

                PdfWriter pw = PdfWriter.GetInstance(doc, fs);
                PdfWriterEvents writerEvent = new PdfWriterEvents("");
                pw.PageEvent = writerEvent;


                doc.Open();
                doc.AddTitle("Statistical report");//文件標題
                doc.AddAuthor("NoMan");//文件作者

                //save image to pdf
                string[] files = Directory.GetFiles(chartTempPath);
                foreach (string file in files)
                {
                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(file);
                    image.ScaleToFit(500, 350);
                    doc.Add(image);//加入影像

                }

                //next page
                doc.NewPage();

                //save tables to pdf
                foreach (KeyValuePair<string, PdfPTable> item in pdfTables)
                {
                    doc.Add(item.Value);
                    doc.Add(new Paragraph(" "));//NewLine
                }
            }
            catch (Exception ex)
            {
                Log.Error($"save2Pdf fail. ErrorMsg={ex.Message}. Stacktrace={ex.StackTrace}");
                doc.Close();
                File.Delete(outputPath);
            }
            finally
            {
                doc.Close();
            }
        }

        private PdfPTable createTable(string serverName, double[] values)
        {
            // Document doc = new Document(PageSize.A4, 50, 50, 80, 50); // 設定PageSize, Margin, left, right, top, bottom
            // MemoryStream ms = new MemoryStream();
            // PdfWriter pw = PdfWriter.GetInstance(doc, ms);

            ////    字型設定
            // 在PDF檔案內容中要顯示中文，最重要的是字型設定，如果沒有正確設定中文字型，會造成中文無法顯示的問題。
            // 首先設定基本字型：kaiu.ttf 是作業系統系統提供的標楷體字型，IDENTITY_H 是指編碼(The Unicode encoding with horizontal writing)，及是否要將字型嵌入PDF 檔中。
            // 再來針對基本字型做變化，例如Font Size、粗體斜體以及顏色等。當然你也可以採用其他中文字體字型。
            //BaseFont bfChinese = BaseFont.CreateFont("C:\\Windows\\Fonts\\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            iTextSharp.text.Font ChFont = new iTextSharp.text.Font(bfChinese, 12);
            iTextSharp.text.Font ChFont_green = new iTextSharp.text.Font(bfChinese, 40, iTextSharp.text.Font.NORMAL, BaseColor.GREEN);
            iTextSharp.text.Font ChFont_msg = new iTextSharp.text.Font(bfChinese, 12, iTextSharp.text.Font.ITALIC, BaseColor.RED);


            // 產生表格 -- START
            // 建立4個欄位表格之相對寬度
            PdfPTable pt = new PdfPTable(new float[] { 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 });
            // 表格總寬
            pt.TotalWidth = 500f;
            pt.LockedWidth = true;


            // 塞入資料 -- START
            // 設定表頭
            PdfPCell header = new PdfPCell(new Phrase(serverName, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
            header.Colspan = 13;
            header.HorizontalAlignment = Element.ALIGN_CENTER;// 表頭內文置中
            header.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);


            pt.AddCell(header);


            PdfPCell cell1 = new PdfPCell(new Phrase("Months", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);
            pt.AddCell(cell1);


            for (int i = 1; i <= 12; i++)
            {
                PdfPCell cell = new PdfPCell(new Phrase(i.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL)));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                pt.AddCell(cell);
            }

            pt.AddCell("Deduplication Rate");

            for (int i = 0; i < 12; i++)
            {
                PdfPCell cell = new PdfPCell(new Phrase(values[i].ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10f, iTextSharp.text.Font.NORMAL)));

                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                pt.AddCell(cell);
            }


            return pt;


        }


        private void createChart(string serverName, double[] values, string imageName)
        {

            //populate dataset with some demo data..

            DataSet dataSet = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Months", typeof(string));
            dt.Columns.Add("Used", typeof(int));
            DataRow r1 = dt.NewRow();

            double maxValue = values.Max();
            double intervalValue = maxValue / 10.0;

            for (int i = 0; i < 12; i++)
            {
                DataRow r = dt.NewRow();
                r[0] = i + 1;
                r[1] = values[i];
                dt.Rows.Add(r);
            }
            dataSet.Tables.Add(dt);


            //prepare chart control...
            Chart chart = new Chart();
            chart.Titles.Add(serverName);
            chart.DataSource = dataSet.Tables[0];
            chart.Width = 1000;
            chart.Height = 700;
            //create serie...
            Series serie1 = new Series();
            serie1.Name = "Serie1";
            serie1.Color = _ChartSerieColor;//Color.FromArgb(112, 255, 200);
            serie1.BorderColor = Color.FromArgb(164, 164, 164);
            serie1.ChartType = SeriesChartType.Column;
            //serie1.ChartType = SeriesChartType.Line;
            serie1.BorderDashStyle = ChartDashStyle.Solid;
            serie1.BorderWidth = 1;
            serie1.ShadowColor = Color.FromArgb(128, 128, 128);
            serie1.ShadowOffset = 1;
            serie1.IsValueShownAsLabel = true;
            serie1.XValueMember = "Months";
            serie1.YValueMembers = "Used";
            serie1.Font = new System.Drawing.Font("Consolas", 8.0f);
            serie1.BackSecondaryColor = Color.FromArgb(0, 102, 153);
            serie1.LabelForeColor = Color.FromArgb(50, 50, 50);
            //serie1.LabelBackColor = Color.LightYellow;



            chart.Series.Add(serie1);
            //create chartareas...

            ChartArea ca = new ChartArea();
            ca.Name = "ChartArea1";
            ca.BackColor = Color.White;
            ca.BorderColor = Color.FromArgb(26, 59, 105);
            ca.BorderWidth = 0;
            ca.BorderDashStyle = ChartDashStyle.Solid;
            ca.AxisX = new Axis();
            ca.AxisY = new Axis();
            ca.AxisX.Minimum = 0;
            ca.AxisX.Interval = 1;
            ca.AxisX.Maximum = 13;


            ca.AxisY.Maximum = getAxisYMaxValue(maxValue);//maxValue + intervalValue;
                                                          //ca.AxisY.Interval = intervalValue;

            chart.ChartAreas.Add(ca);
            //databind...
            chart.DataBind();

            //save result...
            chart.SaveImage(imageName, ChartImageFormat.Png);



        }

        private double getAxisYMaxValue(double d)
        {
            //double d = 25678952125.56;
            int v1 = Math.Round(d).ToString().Length;
            long v2 = long.Parse("1".PadRight(v1, '0'));

            double v3 = Math.Ceiling(d / v2);

            long v4 = v2 / 10;


            double v5 = v3 * v2 + v4;

            return v5;
        }


        private void parseTableFiles(string path, DateTime date)
        {

            for (int month = 0; month < 12; month++)
            {
                var d = date.AddMonths(0 - month);
                Debug.Print($"===========Table Month={d.Month} ============");
                //string filename = getFileNameBy(path, @"(OpsCenter_Deduplication_Rates_by_Policy_Type_\d{2}_", d.Year, d.Month);
                string filename = getFileNameV2By(path, _TableFileNameKey, d.Year, d.Month);
                if (filename == null) continue;
                regexParseTableFile(Path.Combine(path, filename), month);

            }
        }





        private void parseChartFiles(string path, DateTime date)
        {

            for (int month = 0; month < 12; month++)
            {

                var d = date.AddMonths(0 - month);
                Debug.Print($"===========Chart Month={d.Month} ============");
                //string filename = getFileNameBy(path, @"(OpsCenter_last_month_disk_usage_\d{2}_", d.Year, d.Month);
                string filename = getFileNameV2By(path, _ChartFileNameKey, d.Year, d.Month);
                if (filename == null) continue;
                regexParseChartFile(Path.Combine(path, filename), month);

            }
        }


        private string getFileNameBy(string path, string key, int year, int month)
        {
            string pattern = key + $"{month.ToString("00")}_{year}_.*)";
            string[] files = Directory.GetFiles(path);



            foreach (string file in files)
            {
                Match match = Regex.Match(file, pattern);
                if (match.Success)
                {
                    return match.Value;
                }
            }

            return null;

        }


        private string getFileNameV2By(string path, string key, int year, int month)
        {
            //string folderPath = @"D:\Project\MoneyMoneyMoney\create_pdf_chart_and_table\doc_sample\";
            DirectoryInfo dir = new DirectoryInfo(path);

            string pattern = $"{key}*_{month.ToString("00")}_{year}*";
            FileInfo[] files = dir.GetFiles(pattern, SearchOption.TopDirectoryOnly);
            if (files.Count() > 0)
                return files[0].FullName;

            return null;
        }


        private double regexParseTableFile(string fileName, int month)
        {
            string[] lines = File.ReadAllLines(fileName);



            string pattern = @"^(?<ServerName>[a-zA-Z-\s]*),(?<BBB>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<CCC>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<DupRate>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<EEE>""[\d*,?]*[.\d*]*""|\d*[.\d*]*)$";




            foreach (string line in lines)
            {
                Match match = Regex.Match(line, pattern);
                if (match.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"ServerName={match.Groups["ServerName"].ToString()}, BBB={match.Groups["BBB"].ToString()}, CCC={match.Groups["CCC"].ToString()}, DupRate={match.Groups["DupRate"].ToString()}, EEE={match.Groups["EEE"].ToString()}");

                    string serverName = match.Groups["ServerName"].ToString();

                    if (!tableDataSets.ContainsKey(serverName))
                    {
                        tableDataSets.Add(match.Groups["ServerName"].ToString(), new double[12]);
                    }

                    string temp = match.Groups["DupRate"].ToString().Replace(",", string.Empty).Replace("\"", string.Empty); ;
                    double result = Double.Parse(temp);

                    double[] values = tableDataSets[serverName];
                    values[month] = result;



                }
            }

            return 0.0;




        }

        private double regexParseChartFile(string fileName, int index)
        {
            string[] lines = File.ReadAllLines(fileName);



            string pattern = @"^(?<AAA>[a-zA-Z\d-_\s]*),(?<ServerName>[a-zA-Z\d-_\s]*),(?<CCC>[a-zA-Z\d-_\s]*),(?<DDD>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<DiskPollUsedCapacity>""[\d*,?]*[.\d *]*""|\d*[.\d*]*),(?<FFF>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<GGG>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<HHH>""[\d*,?]*[.\d*]*""|\d*[.\d*]*),(?<III>.*)$";




            foreach (string line in lines)
            {
                Match match = Regex.Match(line, pattern);
                if (match.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"AAA={match.Groups["AAA"].ToString()}, ServerName={match.Groups["ServerName"].ToString()}, CCC={match.Groups["CCC"].ToString()}, DDD={match.Groups["DDD"].ToString()}, DiskPollUsedCapacity={match.Groups["DiskPollUsedCapacity"].ToString()}, FFF={match.Groups["FFF"].ToString()}, GGG={match.Groups["GGG"].ToString()}, HHH={match.Groups["HHH"].ToString()}, III={match.Groups["III"].ToString()}");

                    string temp = match.Groups["DiskPollUsedCapacity"].ToString().Replace(",", string.Empty).Replace("\"", string.Empty); ;
                    double result = Double.Parse(temp);

                    string serverName = match.Groups["ServerName"].ToString();

                    if (!chartDataSets.ContainsKey(serverName))
                    {
                        chartDataSets.Add(serverName, new double[12]);
                    }



                    double[] values = chartDataSets[serverName];
                    values[index] = result;

                }
            }

            return 0.0;




        }


        class PdfWriterEvents : IPdfPageEvent
        {
            string watermarkText = string.Empty;

            PdfContentByte cb;
            PdfTemplate template;
            //public BaseFont bfChinese = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            private BaseFont bfChinese = BaseFont.CreateFont("C:\\Windows\\Fonts\\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            public PdfWriterEvents(string watermark)
            {
                watermarkText = watermark;
            }

            public void OnOpenDocument(PdfWriter writer, Document document)
            {
                cb = writer.DirectContent;
                template = cb.CreateTemplate(50, 50);
            }
            public void OnCloseDocument(PdfWriter writer, Document document) { }
            public void OnStartPage(PdfWriter writer, Document document)
            {

                /*
                float fontSize = 80;
                float xPosition = document.PageSize.Width / 2;//300;
                float yPosition = document.PageSize.Height / 2;
                float angle = 45;


                try
                {
                    PdfContentByte under = writer.DirectContentUnder;

                    //BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
                    BaseFont baseFont = BaseFont.CreateFont("kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    under.BeginText();
                    under.SetColorFill(BaseColor.LIGHT_GRAY);
                    under.SetFontAndSize(baseFont, fontSize);
                    under.ShowTextAligned(PdfContentByte.ALIGN_CENTER, watermarkText, xPosition, yPosition, angle);
                    under.EndText();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.Message);
                }
                */
            }
            public void OnEndPage(PdfWriter writer, Document document)
            {
                float fontSize = 10;
                float xPosition = document.PageSize.Width / 2;//300;
                float yPosition = 20;
                float angle = 0;


                try
                {
                    PdfContentByte under = writer.DirectContentUnder;

                    int pageN = writer.PageNumber;
                    String text = "Page " + pageN.ToString();

                    //BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
                    //BaseFont baseFont = BaseFont.CreateFont("kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    under.BeginText();
                    under.SetColorFill(BaseColor.LIGHT_GRAY);
                    under.SetFontAndSize(bfChinese, fontSize);
                    under.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, xPosition, yPosition, angle);
                    under.EndText();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.Message);
                }
            }
            public void OnParagraph(PdfWriter writer, Document document, float paragraphPosition) { }
            public void OnParagraphEnd(PdfWriter writer, Document document, float paragraphPosition) { }
            public void OnChapter(PdfWriter writer, Document document, float paragraphPosition, Paragraph title) { }
            public void OnChapterEnd(PdfWriter writer, Document document, float paragraphPosition) { }
            public void OnSection(PdfWriter writer, Document document, float paragraphPosition, int depth, Paragraph title) { }
            public void OnSectionEnd(PdfWriter writer, Document document, float paragraphPosition) { }
            public void OnGenericTag(PdfWriter writer, Document document, iTextSharp.text.Rectangle rect, String text) { }


        }
    }


}
