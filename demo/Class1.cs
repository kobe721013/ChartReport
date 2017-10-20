using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RGiesecke.DllExport;
using System.Runtime.InteropServices;
using iTextSharp.text.pdf;
using System.Windows.Forms.DataVisualization.Charting;
using iTextSharp.text;
using System.IO;
using System.Drawing;
using System.Data;

namespace demo
{
    public class MyMath
    {

        /*
        [DllExport("ExportReport", CallingConvention = CallingConvention.Cdecl)]
        public static bool TestExport(int month)
        {
           // _imagePath = inputPath;
           // _pdfPath = outputPath;
            Start(month);
            return true;
        }

        */

        [DllExport("Add", CallingConvention = CallingConvention.Cdecl)]
        public static int add(int a, int b)
        {
            return a + b;
        }


        


        static string _imagePath = @"D:\\myChart.png";
        static string _pdfPath = @"D:\\myChart_pdf.pdf";
        static void Start(int month)
        {
            //createFakeChart();

            //save2pdfFormat(false);

            //createFakeTable(true);
        }

        
        static void createFakeTable(bool closeDoc)
        {
           // Document doc = new Document(PageSize.A4, 50, 50, 80, 50); // 設定PageSize, Margin, left, right, top, bottom
           // MemoryStream ms = new MemoryStream();
           // PdfWriter pw = PdfWriter.GetInstance(doc, ms);

            ////    字型設定
            // 在PDF檔案內容中要顯示中文，最重要的是字型設定，如果沒有正確設定中文字型，會造成中文無法顯示的問題。
            // 首先設定基本字型：kaiu.ttf 是作業系統系統提供的標楷體字型，IDENTITY_H 是指編碼(The Unicode encoding with horizontal writing)，及是否要將字型嵌入PDF 檔中。
            // 再來針對基本字型做變化，例如Font Size、粗體斜體以及顏色等。當然你也可以採用其他中文字體字型。
            BaseFont bfChinese = BaseFont.CreateFont("C:\\Windows\\Fonts\\kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
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
            PdfPCell header = new PdfPCell(new Phrase("Server Name", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
            header.Colspan = 13;
            header.HorizontalAlignment = Element.ALIGN_CENTER;// 表頭內文置中
            header.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);


            pt.AddCell(header);


            PdfPCell cell1 = new PdfPCell(new Phrase("Months", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);
            pt.AddCell(cell1);


            for (int i = 1; i <= 12; i++)
            {
                PdfPCell cell = new PdfPCell(new Phrase(i.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(153, 217, 234);
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                pt.AddCell(cell);
            }

            pt.AddCell("Deduplication Rate");
            
            for (int i = 1; i <= 12; i++)
            {
                PdfPCell cell = new PdfPCell(new Phrase((i * 10).ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12f, iTextSharp.text.Font.NORMAL)));
                
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                pt.AddCell(cell);
            }

            
            doc.Add(pt);
            // 塞入資料 -- END
            if(closeDoc)
                doc.Close();
        }


        
        static Document doc = new Document(PageSize.A4, 50, 50, 80, 50);
        /*
        static void save2pdfFormat(bool closeDoc)
        {
            
            try
            {
                FileStream fs = new FileStream(_pdfPath, FileMode.Create);
                PdfWriter.GetInstance(doc, fs);
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(_imagePath);
                doc.Open();
                doc.Add(image);//加入影像
                doc.AddTitle("Tutorial-Add image files");//文件標題
                doc.AddAuthor("einboch");//文件作者
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (doc.IsOpen() && closeDoc) doc.Close();
            }
        }

        static private void createFakeChart()
        {

            //populate dataset with some demo data..
            
            DataSet dataSet = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Months", typeof(string));
            dt.Columns.Add("Used", typeof(int));
            DataRow r1 = dt.NewRow();


            for (int i = 1; i <= 12; i++)
            {
                DataRow r = dt.NewRow();
                r[0] = $"{i}";
                r[1] = 1000+i;
                dt.Rows.Add(r);
            }
            dataSet.Tables.Add(dt);
            
           
            //prepare chart control...
            Chart chart = new Chart();
            chart.Titles.Add($"Server Name");
            chart.DataSource = dataSet.Tables[0];
            chart.Width = 500;
            chart.Height = 350;
            //create serie...
            Series serie1 = new Series();
            serie1.Name = "Serie1";
            serie1.Color = Color.FromArgb(112, 255, 200);
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
            serie1.Font = new System.Drawing.Font("Tahoma", 8.0f);
            serie1.BackSecondaryColor = Color.FromArgb(0, 102, 153);
            serie1.LabelForeColor = Color.FromArgb(100, 100, 100);

            

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

            //ca.AxisY.Minimum = 500;
            //ca.AxisY.Maximum = 1500;

            chart.ChartAreas.Add(ca);
            //databind...
            chart.DataBind();
            
            //save result...
            chart.SaveImage(@"D:\\myChart.png", ChartImageFormat.Png);

            

        }


        static private void test2() {
            Chart chart1 = new Chart();
            ChartArea ca = new ChartArea();
            ca.Name = "ChartArea1";
            ca.BackColor = Color.White;
            ca.BorderColor = Color.FromArgb(26, 59, 105);
            ca.BorderWidth = 0;
            ca.BorderDashStyle = ChartDashStyle.Solid;
            ca.AxisX = new Axis();
            ca.AxisY = new Axis();

            chart1.ChartAreas.Add(ca);

            //---------------------
            chart1.Series.Clear();  //每次使用此function前先清除圖表
            Series series1 = new Series("Di0", 500); //初始畫線條(名稱，最大值)
            series1.Color = Color.Blue; //設定線條顏色
            series1.Font = new System.Drawing.Font("新細明體", 10); //設定字型
            series1.ChartType = SeriesChartType.Line; //設定線條種類

            chart1.ChartAreas[0].AxisX.Minimum = 1;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.Maximum = 12;

            chart1.ChartAreas[0].AxisY.Minimum = 0;//設定Y軸最小值
            chart1.ChartAreas[0].AxisY.Maximum = 500;//設定Y軸最大值
                                                     //chart1.ChartAreas[0].AxisY.Enabled= AxisEnabled.False; //隱藏Y 軸標示
                                                     //chart1.ChartAreas[0].AxisY.MajorGrid.Enabled= true;  //隱藏Y軸標線
            series1.IsValueShownAsLabel = true; //是否把數值顯示在線上

            //把值加入X 軸Y 軸
            for (int i = 1; i <= 12; i++)
            {
                series1.Points.AddXY(i.ToString(), 200+i*10);
            }
           
            chart1.Series.Add(series1);//將線畫在圖上

            chart1.SaveImage(@"D:\\myChart_2.png", ChartImageFormat.Png);
        }
        */
    }
}
