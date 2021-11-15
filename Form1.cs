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
        public double[] bbb;
        public double[] array1;
        public double[] array2;
        public double[] array3;
        public double[] array4;
        public double[] array5;

        public string[,] list;

        List<Thread> threads = new List<Thread>();

        public bool locker = false;

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
            array1 = new double[values.Count];
            array2 = new double[values.Count];
            array3 = new double[values.Count];
            array4 = new double[values.Count];
            array5 = new double[values.Count];

            for (int i = 0; i < values.Count; i++)
            {
                double val0 = Convert.ToDouble(values[i][0]);
                array1[i] = val0;
                array2[i] = val0;
                array3[i] = val0;
                array4[i] = val0;
                array5[i] = val0;
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





        

        #region UserInteraction
        private void button1_Click(object sender, EventArgs e)
        {
            if (locker == false)
            {
                if (array1 != null || array2 != null || array3 != null || array4 != null || array5 != null)
                {
                    if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == false && checkBox4.Checked == false && checkBox5.Checked == false)
                    {
                        MessageBox.Show("Выберите метод сортировки!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (checkBox1.Checked == true)
                        {
                            chart1.Series[0].Points.DataBindY(array1);
                            Thread bubble = new Thread(new ParameterizedThreadStart(BubbleSort));
                            threads.Add(bubble);
                            bubble.Start(array1);
                        }
                        if (checkBox2.Checked == true)
                        {
                            Thread Insert = new Thread(new ParameterizedThreadStart(InsertSort));
                            threads.Add(Insert);
                            Insert.Start(array2);
                        }
                        if (checkBox3.Checked == true)
                        {
                            Thread Shaker = new Thread(new ThreadStart(ShakerSort));
                            threads.Add(Shaker);
                            Shaker.Start();
                        }
                        if (checkBox4.Checked == true)
                        {
                            Thread Quick = new Thread(new ParameterizedThreadStart(QuickSort));
                            threads.Add(Quick);
                            Quick.Start(array4);
                        }
                        if (checkBox5.Checked == true)
                        {
                            Thread Bogo = new Thread(new ParameterizedThreadStart(BogoSort));
                            threads.Add(Bogo);
                            Bogo.Start(array5);
                        }
                        locker = true;
                    }
                }
                else
                {
                    MessageBox.Show("Массив не был загружен. Загрузите массив и повторите попытку!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Нажмите Очистить и повторите заново", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ExportExcel();
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    array1[i] = Convert.ToDouble(list[i, 0]);
                    dataGridView1.Rows.Add(array1[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        async private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                var serviceValues = GetSheetsService().Spreadsheets.Values;
                await ReadAsync(serviceValues);
                for (int i = 0; i < array1.GetLength(0); i++)
                {
                    dataGridView1.Rows.Add(array1[i]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (var item in threads)
            {
                if (item.ThreadState != System.Threading.ThreadState.Stopped)
                    item.Resume();
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (locker == false)
            {
                if (array1 != null || array2 != null || array3 != null || array4 != null || array5 != null)
                {
                    if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == false && checkBox4.Checked == false && checkBox5.Checked == false)
                    {
                        MessageBox.Show("Выберите метод сортировки!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        if (checkBox1.Checked == true)
                        {
                            chart1.Series[0].Points.DataBindY(array1);
                            Thread Inversebubble = new Thread(new ParameterizedThreadStart(InverseBubbleSort));
                            threads.Add(Inversebubble);
                            Inversebubble.Start(array1);
                        }
                        if (checkBox2.Checked == true)
                        {
                            Thread InverseInsert = new Thread(new ParameterizedThreadStart(InverseInsertSort));
                            threads.Add(InverseInsert);
                            InverseInsert.Start(array2);
                        }
                        if (checkBox3.Checked == true)
                        {
                            Thread InverseShaker = new Thread(new ThreadStart(InverseShakerSort));
                            threads.Add(InverseShaker);
                            InverseShaker.Start();
                        }
                        if (checkBox4.Checked == true)
                        {
                            Thread InverseQuick = new Thread(new ParameterizedThreadStart(InverseQuickSort));
                            threads.Add(InverseQuick);
                            InverseQuick.Start(array4);
                        }
                        if (checkBox5.Checked == true)
                        {
                            Thread InverseBogo = new Thread(new ParameterizedThreadStart(InverseBogoSort));
                            threads.Add(InverseBogo);
                            InverseBogo.Start(array5);
                        }
                        locker = true;
                    }
                }
                else
                {
                    MessageBox.Show("Массив не был загружен. Загрузите массив и повторите попытку!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Нажмите Очистить и повторите заново", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (var item in threads)
            {
                if (item.ThreadState == System.Threading.ThreadState.Suspended)
                {
                    item.Resume();
                    item.Abort();
                }
                else
                    item.Abort();
            }
            threads.Clear();
            dataGridView1.Rows.Clear();
            array1 = null;
            array2 = null;
            array3 = null;
            array4 = null;
            array5 = null;
            locker = false;
            chart1.Series[0].Points.Clear();
            chart2.Series[0].Points.Clear();
            chart3.Series[0].Points.Clear();
            chart4.Series[0].Points.Clear();
            chart5.Series[0].Points.Clear();
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            label1.Text = "";
            label2.Text = "";
            label3.Text = "";
            label4.Text = "";
            label5.Text = "";
        }

        #endregion

        #region Sorts

        public void BubbleSort(object arr1)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            double temp;
            for (int i = 0; i < array1.GetLength(0); ++i)
            {
                for (int j = i + 1; j < array1.GetLength(0); ++j)
                {
                    if (array1[i] > array1[j])
                    {
                        temp = array1[i];
                        array1[i] = array1[j];
                        array1[j] = temp;
                        Action action = () => chart1.Series[0].Points.DataBindY(array1);
                        Invoke(action);
                        Thread.Sleep(100);
                    }
                }
            }
            sw.Stop();
            Action action5 = () => label2.Text = sw.ElapsedMilliseconds.ToString()+"ms";
            Invoke(action5);
        }

        public void InsertSort(object arr1)
        {

            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int i = 1; i< array2.GetLength(0); ++i)
            {
                double key = array2[i];
                int j = i - 1;

                while (j >= 0 && array2[j] > key)
                {
                    array2[j + 1] = array2[j];
                    --j;
                }
                 array2[j + 1] = key;
                Action action = () => chart2.Series[0].Points.DataBindY(array2);
                Invoke(action);
                Thread.Sleep(100);
            }
            sw.Stop();
            Action action5 = () => label3.Text = sw.ElapsedMilliseconds.ToString()+"ms";
            Invoke(action5);
        }


        public void BogoSort(object arr1)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            while (!IsSorted())
            {
                RandomSwap();
                Action action = () => chart5.Series[0].Points.DataBindY(array5);
                Invoke(action);
                Thread.Sleep(100);
            }
            sw.Stop();
            Action action5 = () => label4.Text = sw.ElapsedMilliseconds.ToString()+"ms";
            Invoke(action5);
        }
        public bool IsSorted()
        {
            for (int i = 0; i < array5.GetLength(0) - 1; i++)
            {
                if (array5[i] > array5[i + 1])
                {
                    return false;
                }
            }

            return true;
        }

        public void RandomSwap()
        {   
            Random random = new Random();
            var n = array5.GetLength(0);
            while (n > 1)
            {
                --n;
                var i = random.Next(n + 1);
                var temp = array5[i];
                array5[i] = array5[n];
                array5[n] = temp;
            }
        }

        public int Wall(double[] array, int minIndex, int maxIndex)
        {
            var pivot = minIndex - 1;
            for (var i = minIndex; i < maxIndex; ++i)
            {
                if (array[i] < array[maxIndex])
                {
                    ++pivot;
                    Swap(ref array[pivot], ref array[i]);
                }
            }
            ++pivot;
            Swap(ref array[pivot], ref array[maxIndex]);
            Thread.Sleep(100);
            return pivot;
        }

        public double[] QuickSort(double[] array, int minIndex, int maxIndex)
        {

            if (minIndex >= maxIndex)
            {
                return array;
            }

            var pivotIndex = Wall(array, minIndex, maxIndex);
            Action action = () => chart4.Series[0].Points.DataBindY(array4);
            Invoke(action);
            QuickSort(array, minIndex, pivotIndex - 1);
            Invoke(action);
            QuickSort(array, pivotIndex + 1, maxIndex);
            Invoke(action);
            return array;
        }

        public void QuickSort(object a)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            double[] array = (double[])a;
            QuickSort(array, 0, array.Length - 1);
            sw.Stop();
            Action action5 = () => label5.Text = sw.ElapsedMilliseconds.ToString() + "ms";
            Invoke(action5);
        }

        public void ShakerSort()
        {
            for (var i = 0; i < array3.GetLength(0) / 2; i++)
            {
                Action action = () => chart3.Series[0].Points.DataBindY(array3);
                Invoke(action);
                var swapFlag = false;
                //проход слева направо
                for (var j = i; j < array3.GetLength(0) - i - 1; j++)
                {
                    if (array3[j] > array3[j + 1])
                    {
                        Swap(ref array3[j], ref array3[j + 1]);
                        swapFlag = true;
                    }
                }
                Invoke(action);
                //проход справа налево
                for (var j = array3.Length - 2 - i; j > i; j--)
                {
                    if (array3[j - 1] > array3[j])
                    {
                        Swap(ref array3[j - 1], ref array3[j]);
                        swapFlag = true;
                    }
                }
                Invoke(action);
                //если обменов не было выходим
                if (!swapFlag)
                {
                    break;
                }
                Thread.Sleep(100);
            }
        }

        public void Swap(ref double x, ref double y)
        {
            var temp = x;
            x = y;
            y = temp;
        }

        #endregion

        #region InverseSorts
        public void InverseBubbleSort(object arr1)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            double temp;
            for (int i = 0; i < array1.GetLength(0); ++i)
            {
                for (int j = i + 1; j < array1.GetLength(0); ++j)
                {
                    if (array1[i] < array1[j])
                    {
                        temp = array1[i];
                        array1[i] = array1[j];
                        array1[j] = temp;
                        Action action = () => chart1.Series[0].Points.DataBindY(array1);
                        Invoke(action);
                        Thread.Sleep(100);
                    }
                }
            }
            sw.Stop();
            Action action5 = () => label2.Text = sw.ElapsedMilliseconds.ToString() + "ms";
            Invoke(action5);
        }

        public void InverseInsertSort(object arr1)
        {

            Stopwatch sw = new Stopwatch();
            sw.Start();

            for (int i = 1; i < array2.GetLength(0); ++i)
            {
                double key = array2[i];
                int j = i - 1;

                while (j >= 0 && array2[j] < key)
                {
                    array2[j + 1] = array2[j];
                    --j;
                }
                array2[j + 1] = key;
                Action action = () => chart2.Series[0].Points.DataBindY(array2);
                Invoke(action);
                Thread.Sleep(100);
            }
            sw.Stop();
            Action action5 = () => label3.Text = sw.ElapsedMilliseconds.ToString() + "ms";
            Invoke(action5);
        }

        public void InverseBogoSort(object arr1)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            while (!InverseIsSorted())
            {
                RandomSwap();
                Action action = () => chart5.Series[0].Points.DataBindY(array5);
                Invoke(action);
                Thread.Sleep(100);
            }
            sw.Stop();
            Action action5 = () => label4.Text = sw.ElapsedMilliseconds.ToString() + "ms";
            Invoke(action5);
        }
        public bool InverseIsSorted()
        {
            for (int i = 0; i < array5.GetLength(0) - 1; i++)
            {
                if (array5[i] < array5[i + 1])
                {
                    return false;
                }
            }

            return true;
        }

        public void InverseShakerSort()
        {
            for (var i = 0; i < array3.GetLength(0) / 2; i++)
            {
                Action action = () => chart3.Series[0].Points.DataBindY(array3);
                Invoke(action);
                var swapFlag = false;
                //проход слева направо
                for (var j = i; j < array3.GetLength(0) - i - 1; j++)
                {
                    if (array3[j] < array3[j + 1])
                    {
                        Swap(ref array3[j], ref array3[j + 1]);
                        swapFlag = true;
                    }
                }
                Invoke(action);
                //проход справа налево
                for (var j = array3.Length - 2 - i; j > i; j--)
                {
                    if (array3[j - 1] < array3[j])
                    {
                        Swap(ref array3[j - 1], ref array3[j]);
                        swapFlag = true;
                    }
                }
                Invoke(action);
                //если обменов не было выходим
                if (!swapFlag)
                {
                    break;
                }
                Thread.Sleep(100);
            }
        }

        public int InverseWall(double[] array, int minIndex, int maxIndex)
        {
            var pivot = minIndex - 1;
            for (var i = minIndex; i < maxIndex; ++i)
            {
                if (array[i] > array[maxIndex])
                {
                    ++pivot;
                    Swap(ref array[pivot], ref array[i]);
                }
            }
            ++pivot;
            Swap(ref array[pivot], ref array[maxIndex]);
            Thread.Sleep(100);
            return pivot;
        }

        public double[] InverseQuickSort(double[] array, int minIndex, int maxIndex)
        {

            if (minIndex >= maxIndex)
            {
                return array;
            }

            var pivotIndex = InverseWall(array, minIndex, maxIndex);
            Action action = () => chart4.Series[0].Points.DataBindY(array4);
            Invoke(action);
            InverseQuickSort(array, minIndex, pivotIndex - 1);
            Invoke(action);
            InverseQuickSort(array, pivotIndex + 1, maxIndex);
            Invoke(action);
            return array;
        }

        public void InverseQuickSort(object a)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            double[] array = (double[])a;
            InverseQuickSort(array, 0, array.Length - 1);
            sw.Stop();
            Action action5 = () => label5.Text = sw.ElapsedMilliseconds.ToString() + "ms";
            Invoke(action5);
        }
        #endregion
    }
}
