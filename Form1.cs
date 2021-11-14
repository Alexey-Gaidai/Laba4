using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
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

        public static List<double> array1 = new List<double>();
        public static List<double> array2 = new List<double>();
        public static List<double> array3 = new List<double>();
        public static List<double> array4 = new List<double>();
        public static List<double> array5 = new List<double>();

        public string[,] list;

        List<Thread> threads = new List<Thread>();

        public bool pause;

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
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values;
            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return;
            }
            for (int i = 0; i < values.Count; i++)
            {
                double val0 = Convert.ToDouble(values[i][0]);
                array1.Add(val0);
                array2.Add(val0);
                array3.Add(val0);
                array4.Add(val0);
                array5.Add(val0);
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
                for (int i = 0; i < array1.Count; i++)
                {
                    dataGridView1.Rows.Add(array1[i]);
                }
                //BuildChart(0, steps.Count+1, 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ExportExcel();
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    array1[i] = Convert.ToDouble(list[i,0]);
                    dataGridView1.Rows.Add(array1[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pause = true;
            if(checkBox1.Checked == true)
            {
                Thread bubble = new Thread(new ParameterizedThreadStart(BubbleSort));
                threads.Add(bubble);
                bubble.Start(array1);
            }
            if(checkBox2.Checked == true)
            {
                Thread Insert = new Thread(new ParameterizedThreadStart(InsertSort));
                threads.Add(Insert);
                Insert.Start(array2);
            }
            if(checkBox4.Checked == true)
            {
                double[] arr = new double[array4.Count()];
                for (int i = 0; i < array4.Count(); i++)
                {
                    arr[i] = array4[i];
                }
                Thread Quick = new Thread(new ParameterizedThreadStart(QuickSort));
                threads.Add(Quick);
                Quick.Start(arr);
            }
            if(checkBox5.Checked == true)
            {
                Thread Bogo = new Thread(new ParameterizedThreadStart(BogoSort));
                threads.Add(Bogo);
                Bogo.Start(array5);
            }

        }

        public void addchart(Chart chart)
        {
            groupBox1.Controls.Add(chart);
        }
        

        public void BubbleSort(object arr1)
        {
            Chart chart1 = new Chart();
            ChartArea area = new ChartArea();
            chart1.Size = new System.Drawing.Size(290, 210);
            chart1.Location = new System.Drawing.Point(6, 19);
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = array1.Count()+1;
            area.AxisX.MajorGrid.Enabled = true;
            chart1.ChartAreas.Add(area);
            Series series1 = new Series("Сортировка пузырьком");
            series1.ChartType = SeriesChartType.Column;
            series1.Legend = "Legend1";
            chart1.Series.Add(series1);
            for (int i = 0; i < array1.Count(); i++)
                chart1.Series[0].Points.Add(array1[i]);
            Action action2 = () => chart1.Update();
            Invoke(action2);
            Action action = () => addchart(chart1);
            Invoke(action);
            double temp;
            Thread.Sleep(500);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            for (int i = 0; i < array1.Count(); ++i)
            {
                for (int j = i + 1; j < array1.Count(); ++j)
                {
                    if (array1[i] > array1[j])
                    {
                        temp = array1[i];
                        array1[i] = array1[j];
                        array1[j] = temp;
                        Action action3 = () => chart1.Series[0].Points.Clear();
                        Invoke(action3);
                        for (int k = 0; k < array1.Count(); k++)
                        {
                            Action action4 = () => chart1.Series[0].Points.Add(array1[k]);
                            Invoke(action4);
                        }
                        Thread.Sleep(500);
                        if (pause == false)
                            Thread.Sleep(10000000);
                    }
                }
            }
            sw.Stop();
            double aaa = sw.ElapsedMilliseconds / 1000;
            Action action5 = () => label2.Text = aaa.ToString();
            Invoke(action5);
            //return arr;
        }

        public void InsertSort(object arr1)
        {
            Chart chart2 = new Chart();
            ChartArea area = new ChartArea();
            chart2.Size = new System.Drawing.Size(290, 210);
            chart2.Location = new System.Drawing.Point(298, 19);
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = array2.Count()+1;
            area.AxisX.MajorGrid.Enabled = true;
            chart2.ChartAreas.Add(area);
            Series series1 = new Series("Сортировка пузырьком");
            series1.ChartType = SeriesChartType.Column;
            series1.Legend = "Legend1";
            chart2.Series.Add(series1);
            for (int i = 0; i < array2.Count(); i++)
                chart2.Series[0].Points.Add(array2[i]);
            Action action = () => addchart(chart2);
            Invoke(action);
            Action action2 = () => chart2.Update();
            Invoke(action2);
            Thread.Sleep(500);

            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int i = 1; i< array2.Count(); ++i)
            {
                double key = array2[i];
                int j = i - 1;

                while (j >= 0 && array2[j] > key)
                {
                    array2[j + 1] = array2[j];
                    --j;
                }
                 array2[j + 1] = key;
                Action action3 = () => chart2.Series[0].Points.Clear();
                Invoke(action3);
                for (int k = 0; k < array2.Count(); k++)
                {
                    Action action4 = () => chart2.Series[0].Points.Add(array2[k]);
                    Invoke(action4);
                }
                Thread.Sleep(500);
                if (pause == false)
                    Thread.Sleep(10000000);
            }
            sw.Stop();
            double aaa = sw.ElapsedMilliseconds / 1000;
            Action action5 = () => label3.Text = aaa.ToString();
            Invoke(action5);
            //return arr;
        }


        public void BogoSort(object arr1)
        {
            Chart chart2 = new Chart();
            ChartArea area = new ChartArea();
            chart2.Size = new System.Drawing.Size(290, 210);
            chart2.Location = new System.Drawing.Point(8, 229);
            area.AxisX.Minimum = 0;
            area.AxisX.Maximum = array3.Count()+1;
            area.AxisX.MajorGrid.Enabled = true;
            chart2.ChartAreas.Add(area);
            Series series1 = new Series("Сортировка пузырьком");
            series1.ChartType = SeriesChartType.Column;
            series1.Legend = "Legend1";
            chart2.Series.Add(series1);
            for (int i = 0; i < array3.Count(); i++)
                chart2.Series[0].Points.Add(array3[i]);
            Action action = () => addchart(chart2);
            Invoke(action);
            Action action2 = () => chart2.Update();
            Invoke(action2);
            Thread.Sleep(500);

            Stopwatch sw = new Stopwatch();
            sw.Start();

            while (!IsSorted())
            {
                RandomSwap();
                Action action3 = () => chart2.Series[0].Points.Clear();
                Invoke(action3);
                for (int k = 0; k < array3.Count(); k++)
                {
                    Action action4 = () => chart2.Series[0].Points.Add(array3[k]);
                    Invoke(action4);
                }
                Thread.Sleep(500);
            }
            sw.Stop();
            double aaa = sw.ElapsedMilliseconds / 1000;
            Action action5 = () => label4.Text = aaa.ToString();
            Invoke(action5);
        }
        public static bool IsSorted()
        {
            for (int i = 0; i < array3.Count - 1; i++)
            {
                if (array3[i] > array3[i + 1])
                {
                    return false;
                }
            }

            return true;
        }

        public static void RandomSwap()
        {   
            Random random = new Random();
            var n = array3.Count();
            while (n > 1)
            {
                --n;
                var i = random.Next(n + 1);
                var temp = array3[i];
                array3[i] = array3[n];
                array3[n] = temp;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            foreach (var item in threads)
            {
                if (item.ThreadState != System.Threading.ThreadState.Stopped)
                    item.Suspend();
            }
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (var item in threads)
            {
                item.Abort();
            }
            groupBox1.Controls.Clear();
            threads.Clear();
            dataGridView1.Rows.Clear();
            array1.Clear();
            array2.Clear();
            array3.Clear();
            array4.Clear();
            array5.Clear();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (var item in threads)
            {
                if (item.ThreadState != System.Threading.ThreadState.Stopped)
                    item.Resume();
            }
            
        }
        public int Wall(double[] array, int minIndex, int maxIndex)
        {

            var pivot = minIndex - 1;
            for (var i = minIndex; i < maxIndex; ++i)
            {
                Action action4 = () => chart1.Series[0].Points.Clear();
                Invoke(action4);
                if (array[i] < array[maxIndex])
                {
                    ++pivot;
                    Swap(ref array[pivot], ref array[i]);
                    for (int g = 0; g < array.GetLength(0); g++)
                    {
                        Action action3 = () => chart1.Series[0].Points.Add(array[g]);
                        Invoke(action3);
                    }
                    Thread.Sleep(1000);
                }
            }

            ++pivot;
            Swap(ref array[pivot], ref array[maxIndex]);
            Action action5 = () => chart1.Series[0].Points.Clear();
            Invoke(action5);
            for (int g = 0; g < array.GetLength(0); g++)
            {
                Action action3 = () => chart1.Series[0].Points.Add(array[g]);
                Invoke(action3);
            }
            Thread.Sleep(1000);
            return pivot;
        }

        public double[] QuickSort(double[] array, int minIndex, int maxIndex)
        {

            if (minIndex >= maxIndex)
            {
                return array;
            }

            var pivotIndex = Wall(array, minIndex, maxIndex);

            QuickSort(array, minIndex, pivotIndex - 1);
            QuickSort(array, pivotIndex + 1, maxIndex);

            return array;
        }

        public void QuickSort(object a)
        {
            double[] array = (double[])a;
            for (int i = 0; i < array.GetLength(0); i++)
            {
                Action action3 = () => chart1.Series[0].Points.Add(array[i]);
                Invoke(action3);
            }
            QuickSort(array, 0, array.Length - 1);
            Thread.Sleep(1000);
        }

        public void Swap(ref double x, ref double y)
        {
            var temp = x;
            x = y;
            y = temp;
        }
    }
}
