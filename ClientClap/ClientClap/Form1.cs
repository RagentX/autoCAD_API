using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.IO;
using System.Threading.Tasks;

namespace ClientClap
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = openFileDialog1.FileName;
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            
            
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            //string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
            int asme = Convert.ToInt32(ObjWorkSheet.Cells[26,9].Text.ToString());//считываем текст в строку
            int DN = Convert.ToInt32(ObjWorkSheet.Cells[26, 11].Text.ToString());//считываем текст в строку
            string form = ObjWorkSheet.Cells[26, 10].Text.ToString();//считываем текст в строку
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой
            int[] values = { 20, 25, 40, 50, 80, 100, 150, 200 };
            string[] parClap = getData(Array.IndexOf(values,DN),form.ToLower(),asme);
            //MessageBox.Show($"https://localhost:44394/Index?klap_par={parClap[0]},{parClap[1]},{parClap[2]}&klap=Клапан");
            


            string s = await RequestAsync(parClap);
            MessageBox.Show(s);

        }
        private async Task<String> RequestAsync(string[] parClap)
        {
            string url = $"https://localhost:44394/Index?klap_par={parClap[0]},{parClap[1]},{parClap[2]}&klap=Клапан";
            WebRequest request = WebRequest.Create(url);
            WebResponse response = await request.GetResponseAsync().ConfigureAwait(true);
            Stream stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);
            return await reader.ReadToEndAsync();
        }


        public static string[][] getDataFromCSV(string pathCsvFile)
        {
            List<string[]> data = new List<string[]>();
            System.IO.StreamReader file = new System.IO.StreamReader(pathCsvFile);
            string line;
            while ((line = file.ReadLine()) != null)
            {
                String[] parts_of_line = line.Split(';');
                string[] mass = new string[parts_of_line.Length];
                for (int i = 0; i < parts_of_line.Length; i++)
                {
                    parts_of_line[i] = parts_of_line[i].Trim();
                    mass[i] = parts_of_line[i];
                }
                data.Add(mass);
            }

            return data.ToArray();
        }


        public static string[] getData(int value, string form, int class_asme)
        {
            string[][] dataMassCsvFileA = getDataFromCSV(@"C:\a\bred\A.csv");
            string[][] dataMassCsvFileB = getDataFromCSV(@"C:\a\bred\B.csv");
            string[][] dataMassCsvFileC = getDataFromCSV(@"C:\a\bred\C.csv");


            for (int i = 0; i < dataMassCsvFileA.GetLength(0); i++)
            {
                for (int j = 0; j < dataMassCsvFileA[i].GetLength(0); j++)
                {
                    Console.Write(dataMassCsvFileA[i][j]);
                    Console.Write('\t');

                }
                Console.Write("\n");
            }
            Console.Write("\n\n\n");

            int A = 0;
            for (int i = 0; i < dataMassCsvFileA[0].GetLength(0); i++)
            {
                if (dataMassCsvFileA[0][i] == class_asme.ToString() && dataMassCsvFileA[1][i] == form)
                {
                    A = Int32.Parse(dataMassCsvFileA[value + 2][i]);
                    break;
                }
            }

            int B = 0;
            for (int i = 0; i < dataMassCsvFileB[0].GetLength(0); i++)
            {
                if (dataMassCsvFileB[0][i] == class_asme.ToString())
                {
                    B = Int32.Parse(dataMassCsvFileB[value + 2][i]);
                    break;
                }
            }

            int C = 0;
            for (int i = 0; i < dataMassCsvFileC[0].GetLength(0); i++)
            {
                if (dataMassCsvFileC[0][i] == class_asme.ToString())
                {
                    C = Int32.Parse(dataMassCsvFileC[value + 1][i]);
                    break;
                }
            }
            return new string[] { A.ToString(), B.ToString(), C.ToString() };
        }
        /*
        private async void button1_Click(object sender, EventArgs e)
        {
            string s = await RequestAsync();
            MessageBox.Show(s);
        }
        private async Task<String> RequestAsync()
        {
            string rez = "asf";
            string url = "https://localhost:44394/Index?klap_par=25,35,15&klap=Клапан";
            WebRequest request = WebRequest.Create(url);
            WebResponse response = request.GetResponse();
            Stream stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);
            response.Close();
            rez = reader.ReadToEnd();
            return rez;
        }*/
    }
    
}
