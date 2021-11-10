using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;

namespace Laba4
{
    public partial class Form1 : Form
    {
        public static List<double> steps = new List<double>();//список точек
        public double[] sss;
        public string[,] list;

        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1Kcvpqi-I6wY0HSFGehgdVp_tS70Fk2KQroZT39Z8S5Q";
        private const string GoogleCredentialsFileName = "google-credentials.json";
        private const string ReadRange = "Лист1!A:A";
        public Form1()
        {
            InitializeComponent();
        }

        private static SheetsService GetSheetsService()//получаем ответ от сервера
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }

        private async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)//выполняем чтение
        {
            var response = await valuesResource.Get(textBox2.Text, ReadRange).ExecuteAsync();
            var values = response.Values;
            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return;
            }
            sss = new double[values.Count];
            for (int i = 0; i < values.Count; i++)
            {
                double val0 = Convert.ToDouble(values[i][0]);
                steps.Add(val0);
                sss[i] = val0;
            }

        }

        private void ExportExcel()//получение точек из эксель
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл базы данных";

            if (!(ofd.ShowDialog() == DialogResult.OK))
                MessageBox.Show("", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;
            if (lastRow <= 2)
            {
                MessageBox.Show("Недостаточное количество точек", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                list = new string[lastRow, lastColumn];

                for (int j = 0; j < 1; j++)
                    for (int i = 0; i < lastRow; i++)
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            ObjWorkExcel.Quit();
            GC.Collect();
        }





        async private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var serviceValues = GetSheetsService().Spreadsheets.Values;
                await ReadAsync(serviceValues);
                for (int i = 0; i < steps.Count; i++)
                {
                    dataGridView1.Rows.Add(sss[i]);
                }
                //BuildChart(0, steps.Count+1, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void BuildChart(double x_min, double x_max, double dx)
        {
            Chart chart1 = new Chart();
            //groupBox1.Controls.Add(chart1);
            ChartArea area = new ChartArea();
            area.AxisX.Minimum = x_min;
            area.AxisX.Maximum = x_max;
            area.AxisX.MajorGrid.Enabled = true;
            chart1.ChartAreas.Add(area);
            Series series1 = new Series();
            series1.ChartType = SeriesChartType.Column;
            chart1.Series.Add(series1);
            for (int i = 0; i < sss.GetLength(0); i++)
                chart1.Series[0].Points.Add(sss[i]);//отрисовка
            chart1.Update();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ExportExcel();
                sss = new double[list.GetLength(0)];
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    sss[i] = Convert.ToDouble(list[i,0]);
                    dataGridView1.Rows.Add(sss[i]);
                }
                BuildChart(0, sss.GetLength(0) + 1, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        async private void button1_Click(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                sss = await method(sss);
            }
            if(checkBox2.Checked == true)
            {
                sss = InsertSort(sss);
                BuildChart(0, sss.GetLength(0) + 1, 1);
            }

        }

        async Task<Double[]> method(double[] arr)//асинхроним расчеты метода
        {
            var result = await Task.Run(() => BubbleSort(arr));
            return result;
        }

        public void addchart(Chart chart1)
        {
            groupBox1.Controls.Add(chart1);
        }

        public double[] BubbleSort(double[] arr)
        {
            Chart chart1 = new Chart();
            
            ChartArea area = new ChartArea();
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = arr.Length+1;
            area.AxisX.MajorGrid.Enabled = true;
            chart1.ChartAreas.Add(area);
            Series series1 = new Series();
            series1.ChartType = SeriesChartType.Column;
            chart1.Series.Add(series1);
            for (int i = 0; i < arr.GetLength(0); i++)
                chart1.Series[0].Points.Add(arr[i]);//отрисовка
            Action action2 = () => chart1.Update();
            Invoke(action2);
            Action action = () => addchart(chart1);//так как этот метод находится в другом потоке то вызываем через инвоук
            Invoke(action);
            double temp;
            Thread.Sleep(3000);

            for (int i = 0; i<arr.Length; ++i)
            {
                for (int j = i + 1; j<arr.Length; ++j)
                {
                    if (arr[i] > arr[j])
                    {
                        temp = arr[i];
                        arr[i] = arr[j];
                        arr[j] = temp;
                        Action action3 = () => chart1.Series[0].Points.Clear();
                        Invoke(action3);
                        for (int k = 0; k < sss.GetLength(0); k++)
                        {
                            Action action4 = () => chart1.Series[0].Points.Add(arr[k]);
                            Invoke(action4);
                        }
                        Thread.Sleep(1000);
                    }
}
            }
            return arr;
        }

        public double[] InsertSort(double[] arr)
        {
            for (int i = 1; i<arr.Length; ++i)
            {
                double key = arr[i];
                int j = i - 1;

                while (j >= 0 && arr[j] > key)
                {
                    arr[j + 1] = arr[j];
                    --j;
                }
                 arr[j + 1] = key;
            }
            return arr;
        }
    }
}
