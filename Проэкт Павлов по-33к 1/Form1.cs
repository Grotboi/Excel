using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Tools = Microsoft.Office.Tools.Excel;
using System.IO;

namespace Построение_диаграмм_в_Excel
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            //Показательные функции
            if (textBox1.Text == "")
            {
                try
                {
                    double osnov = Convert.ToDouble(textBox5.Text);
                    double step = Convert.ToDouble(textBox6.Text);
                    Excel.Application Excel_ = new Excel.Application();
                    Excel_.Visible = false;
                    Excel.Workbook WorkBook_;
                    Excel.Worksheet Sheet_;
                    WorkBook_ = Excel_.Workbooks.Add();
                    Sheet_ = (Excel.Worksheet)WorkBook_.Sheets[1];
                    Sheet_.Cells[1, 1] = "x";
                    Sheet_.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlRight;
                    Sheet_.Cells[1, 2] = "y";
                    Sheet_.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlRight;
                    Sheet_.Cells[2, 1] = osnov;
                    Sheet_.Cells[2, 2] = "=POWER(A2," + step + ")";

                    Excel.Range range = Excel_.Range["A2", "B2"];
                    Excel.Chart chart = Excel_.ActiveWorkbook.Charts.Add(After: Excel_.ActiveSheet);

                    chart.ChartWizard(Source: range, Gallery: Excel.XlChartType.xlLineMarkers, Format: 12, Title: "График показательной функции y = (x)^" + step);
                    chart.ChartStyle = 45;
                    chart.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap, Excel.XlPictureAppearance.xlScreen);
                    if (pictureBox1.Image != null)
                    {
                        pictureBox1.Image.Dispose();
                    }
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    chart.Export(AppDomain.CurrentDomain.BaseDirectory + @"\excel_chart_export.bmp", "BMP", Excel_);
                    pictureBox1.Image = new Bitmap(@"excel_chart_export.bmp");
                    Excel_.DisplayAlerts = false;
                    Excel_.ActiveWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Book1.xls", Excel.XlSaveAsAccessMode.xlNoChange);
                    Excel_.Quit();
                    MessageBox.Show("Excel файл создан, вы можете найти его в корневой директории");
                }
                catch (FormatException)
                {
                    MessageBox.Show("Введите число");
                }
                textBox5.Text = null;
                textBox6.Text = null;
            }
            //Тригонометрические функции
            if (textBox1.Text != "")
            {
                try
                {
                    double nach = Convert.ToDouble(textBox1.Text);
                    double shag = Convert.ToDouble(textBox3.Text);
                    double kon = Convert.ToDouble(textBox2.Text);
                }
                catch (FormatException)
                {
                    MessageBox.Show("Введите число");
                }
                    double nach1 = Convert.ToDouble(textBox1.Text);
                    double shag1 = Convert.ToDouble(textBox3.Text);
                    double kon1 = Convert.ToDouble(textBox2.Text);
                    Excel.Application Excel_ = new Excel.Application();
                    Excel_.Visible = false;
                    Excel.Workbook WorkBook_;
                    Excel.Worksheet Sheet_;
                    Excel.Range excelcells;
                    
                    WorkBook_ = Excel_.Workbooks.Add();
                    Sheet_ = (Excel.Worksheet)WorkBook_.Sheets[1];
                    Sheet_.Cells[1, 1] = shag1;
                    Sheet_.Cells[3, 1] = nach1;
                    double q = (kon1- nach1) / shag1;
                    double x = Math.Abs(q);

                    Sheet_.Cells[3, 2].Formula = "=" + textBox4.Text + "(A3)";

                    if (kon1 < 0)
                    {
                        for (double i = 4; i < 4 + x; i++)
                        {
                            Sheet_.Cells[i, 1] = nach1 - shag1 - (shag1 * (i - 4));
                            Sheet_.Cells[i, 2] = String.Format("=" + textBox4.Text + "(A{0})", i);
                        }
                    }
                    if (kon1 > 0)
                    {
                        for (double i = 4; i < 4 + x; i++)
                        {
                            Sheet_.Cells[i, 1] = nach1 + shag1 + (shag1 * (i - 4));
                            Sheet_.Cells[i, 2] = String.Format("=" + textBox4.Text + "(A{0})", i);
                        }
                    }

                    excelcells = Sheet_.Columns["B", Type.Missing];
                    Excel.Chart chart = Excel_.ActiveWorkbook.Charts.Add(After: Excel_.ActiveSheet);
                    chart.ChartWizard(Source: excelcells, Gallery: Excel.XlChartType.xlLineMarkers, Format: 12, Title: "График функции y = " + textBox4.Text + "(x)");
                    chart.ChartStyle = 45;
                    chart.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap, Excel.XlPictureAppearance.xlScreen);
                    Sheet_.Cells[2, 1] = "x";
                    Sheet_.Cells[2, 1].HorizontalAlignment = Excel.Constants.xlRight;
                    Sheet_.Cells[2, 2] = "y";
                    Sheet_.Cells[2, 2].HorizontalAlignment = Excel.Constants.xlRight;

                    if (pictureBox1.Image != null)
                    {
                        pictureBox1.Image.Dispose();
                    }
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    chart.Export(AppDomain.CurrentDomain.BaseDirectory + @"excel_chart_export.bmp", "BMP", Excel_);
                    pictureBox1.Image = new Bitmap(@"excel_chart_export.bmp");

                    Excel_.DisplayAlerts = false;
                    Excel_.ActiveWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + @"Book1.xls", Excel.XlSaveAsAccessMode.xlNoChange);
                    Excel_.Quit();
                    MessageBox.Show("Excel файл создан, вы можете найти его в корневой директории");

                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;

            }
        }
    }
}
