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

            string[] ram_par = {    "Масса, кг: ", "Заказчик: ", "Потребитель: ", "Установка: ", "Позиция: ",
                                    "Модель привода: ", "Модель клапана: " , "Размер: DN " , "присоединение: PN " , "ХЗ1",
                                    "ХЗ2", "ХЗ3" , "ХЗ4" , "ХЗ5" , "ХЗ6"   };
            
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; 
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int asme = Convert.ToInt32(ObjWorkSheet.Cells[26,9].Text.ToString());
            int DN = Convert.ToInt32(ObjWorkSheet.Cells[26, 11].Text.ToString());
            string form = ObjWorkSheet.Cells[26, 10].Text.ToString();
            string[] priv = ObjWorkSheet.Cells[26, 18].Text.ToString().Split('/');

            ram_par[1] += ObjWorkSheet.Cells[2, 3].Text.ToString();
            ram_par[2] += ObjWorkSheet.Cells[3, 3].Text.ToString();
            ram_par[3] += ObjWorkSheet.Cells[5, 3].Text.ToString();
            ram_par[5] += ObjWorkSheet.Cells[26, 18].Text.ToString();
            ram_par[6] += ObjWorkSheet.Cells[26, 4].Text.ToString();
            ram_par[7] += ObjWorkSheet.Cells[26, 11].Text.ToString();
            ram_par[8] += ObjWorkSheet.Cells[26, 8].Text.ToString();



            ObjWorkBook.Close(false, Type.Missing, Type.Missing); 
            ObjWorkExcel.Quit(); 
            GC.Collect(); 
            int[] values = { 20, 25, 40, 50, 80, 100, 150, 200 };
            string parClap = getData(Array.IndexOf(values,DN),form.ToLower(),asme);
            bool privPar = false;
            if (priv[0] == "88") privPar = true;
            ram_par[0] += (getWeight(Array.IndexOf(values, DN), asme)+getWeightPriv(Convert.ToInt32(priv[1]),false)).ToString();
            //MessageBox.Show(parClap[0]);
            string ram_par_str = "";
            foreach (string i in ram_par)
            {
                ram_par_str += i + ";";
            }
            
            ram_par_str.Trim(';');
            //MessageBox.Show(ram_par_str);
            string s = await RequestAsync(parClap, ram_par_str, getDataPriv(Convert.ToInt32(priv[1]), privPar)) ;
            MessageBox.Show(s);
            

        }
        private async Task<String> RequestAsync(string parClap , string ram_par, string priv_par)
        {
            string url = $"https://localhost:44394/Index?klap_par={parClap}&klap=Клапан&ram_par={ram_par}&priv_par={priv_par}";
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


        public static string getData(int value, string form, int class_asme)
        {
            string[][] dataMassCsvFileA = getDataFromCSV(@"..\..\..\..\csvFiles\A.csv");
            string[][] dataMassCsvFileB = getDataFromCSV(@"..\..\..\..\csvFiles\B.csv");
            string[][] dataMassCsvFileC = getDataFromCSV(@"..\..\..\..\csvFiles\C.csv");


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
            return $"{A.ToString()};{B.ToString()};{C.ToString()}";
        }
        public static string getDataPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            string[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\Priv.csv");
            int parSize = 2;
            if (par) parSize++;
            return  $"∅{dataMassCsvFile[1 + nomberSize][1]};" +
                    $"{dataMassCsvFile[1 + nomberSize][parSize]};" +
                    $"{dataMassCsvFile[1 + nomberSize][4]};" +
                    $"{dataMassCsvFile[1 + nomberSize][5]}";

        }
        public static int getWeightPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            string[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\weightsDimen.csv");
            int parSize = 1;
            if (par) parSize++;
            return Convert.ToInt32( dataMassCsvFile[nomberSize+1][parSize] );

        }
        public static int getWeight(int value, int class_asme)
        {
            string[][] dataMassCsvFileA = getDataFromCSV(@"..\..\..\..\csvFiles\bredWeights.csv");
            for (int i = 0; i < dataMassCsvFileA[0].GetLength(0); i++)
            {
                if (dataMassCsvFileA[0][i] == class_asme.ToString())
                {
                    return (Int32.Parse(dataMassCsvFileA[value + 1][i]));
                }
            }
            return 0;

        }
        private void button2_Click(object sender, EventArgs e)
        {

            label1.Text = getWeightPriv(6,true).ToString();
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
