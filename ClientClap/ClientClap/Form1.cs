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
        List<string> excelPar = new List<string>();
        string filename;
        public Form1()
        {
            InitializeComponent();
        }
        private async void button1_Click(object sender, EventArgs e)
        {



        }
        private async Task<String> RequestAsync(string filename, string hand, string parClap, string ram_par, string priv_par, string backParts, string frontParts)
        {
            StreamReader reader = null;
            try
            {
                string url = $"https://localhost:44394/Index?klap_par={parClap}&klap=Клапан&ram_par={ram_par}&priv_par={priv_par}&"
                    + $"backParts={backParts}&frontParts={frontParts}&hand={hand}&filename={filename}";
                WebRequest request = WebRequest.Create(url);
                WebResponse response = await request.GetResponseAsync().ConfigureAwait(true);
                Stream stream = response.GetResponseStream();
                reader = new StreamReader(stream);
                return await reader.ReadToEndAsync();
            } finally
            {
                reader.Close();
            }
        }


        public static string[][] getDataFromCSV(string pathCsvFile)
        {
            List<string[]> data = new List<string[]>();
            using (System.IO.StreamReader file = new System.IO.StreamReader(pathCsvFile))
            {
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
        }


        //todo remove unused wrapping
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
                Console.Write('\n');
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
            return $"{A};{B};{C}";
        }
        public static string getDataPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            string[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\Priv.csv");
            int parSize = 2;
            if (par) parSize++;
            return $"∅{dataMassCsvFile[1 + nomberSize][1]};" +
                    $"{dataMassCsvFile[1 + nomberSize][parSize]};" +
                    $"{dataMassCsvFile[1 + nomberSize][4]};" +
                    $"∅{dataMassCsvFile[1 + nomberSize][5]}";

        }
        public static int getWeightPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            string[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\weightsDimen.csv");
            int parSize = par ? 2 : 1;
            return Convert.ToInt32(dataMassCsvFile[nomberSize + 1][parSize]);

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
        public static string getFrontPart(string Part)
        {
            Part = Part.Replace("\"", "").Trim().ToLower();
            string[][] dataMassCsvFileParts = getDataFromCSV(@"..\..\..\..\csvFiles\artFrontParts.csv");
            List<String> parts = new List<String>();
            List<String> numberPurts = new List<String>();
            for (int i = 0; i < dataMassCsvFileParts.Length; i++)
            {
                parts.Add(dataMassCsvFileParts[i][0].Replace("\"", "").Trim().ToLower());
                numberPurts.Add(dataMassCsvFileParts[i][1].Trim());
            }
            int a = Array.IndexOf(parts.ToArray(), Part);
            if (a != -1) return numberPurts[a];
            return "0";
        }
        public static string getBackPart(string Part)
        {
            Part = Part.Replace("\"", "").Trim().ToLower();
            string[][] dataMassCsvFileParts = getDataFromCSV(@"..\..\..\..\csvFiles\artBackParts.csv");
            List<String> parts = new List<String>();
            List<String> numberPurts = new List<String>();
            for (int i = 0; i < dataMassCsvFileParts.Length; i++)
            {
                parts.Add(dataMassCsvFileParts[i][0].Replace("\"", "").Trim().ToLower());
                numberPurts.Add(dataMassCsvFileParts[i][1].Trim());
            }
            int a = Array.IndexOf(parts.ToArray(), Part);
            if (a != -1) return numberPurts[a];
            return "0";
        }


        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            listBox1.Items.Clear();
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            filename = openFileDialog1.FileName;
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            int start = 14;
            int fin = start;
            while (true)
            {
                if (ObjWorkSheet.Cells[fin, 2].Text.ToString() == "") break;
                listBox1.Items.Add(ObjWorkSheet.Cells[fin, 2].Text.ToString());
                excelPar.Add("Модель клапана: " + ObjWorkSheet.Cells[fin, 4].Text.ToString() +
                            "\nМодель привода: " + ObjWorkSheet.Cells[fin, 18].Text.ToString() +
                            "\nКласс давления: " + ObjWorkSheet.Cells[fin, 8].Text.ToString() + "PN , " +
                            ObjWorkSheet.Cells[fin, 9].Text.ToString() + " ASME" +
                            "\nDN: " + ObjWorkSheet.Cells[fin, 11].Text.ToString());
                fin++;
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button2.Enabled = true;
            label1.Text = excelPar[listBox1.SelectedIndex];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            create_Arm(ObjWorkSheet, listBox1.SelectedIndex + 14, listBox1.Items[listBox1.SelectedIndex].ToString());
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
        }

        //todo StringBuilder
        public async void create_Arm(Excel.Worksheet ObjWorkSheet, int numberString, string fileNameRezDvg)
        {
            string handString = "Ручной дублер";
            string[] ram_par = {    "Масса, кг: ", "Заказчик: ", "Потребитель: ", "Установка: ", "Позиция: ",
                                    "Модель привода: ", "Модель клапана: " , "Размер: DN " , "присоединение: PN " , "ХЗ1",
                                    "ХЗ2", "ХЗ3" , "ХЗ4" , "ХЗ5" , "ХЗ6"   };

            
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int asme = Convert.ToInt32(ObjWorkSheet.Cells[numberString, 9].Text.ToString());
            int DN = Convert.ToInt32(ObjWorkSheet.Cells[numberString, 11].Text.ToString());
            string form = ObjWorkSheet.Cells[numberString, 10].Text.ToString();
            string[] priv = ObjWorkSheet.Cells[numberString, 18].Text.ToString().Split('/');

            ram_par[1] += ObjWorkSheet.Cells[2, 3].Text.ToString();
            ram_par[2] += ObjWorkSheet.Cells[3, 3].Text.ToString();
            ram_par[3] += ObjWorkSheet.Cells[5, 3].Text.ToString();
            ram_par[4] += ObjWorkSheet.Cells[numberString, 2].Text.ToString();
            ram_par[5] += ObjWorkSheet.Cells[numberString, 18].Text.ToString();
            ram_par[6] += ObjWorkSheet.Cells[numberString, 4].Text.ToString();
            ram_par[7] += ObjWorkSheet.Cells[numberString, 11].Text.ToString();
            ram_par[8] += ObjWorkSheet.Cells[numberString, 8].Text.ToString();
            string[] numberPartsExcel = ObjWorkSheet.Cells[numberString, 28].Text.ToString().Split(',');

            int start = 1;
            while (true)
            {
                if (ObjWorkSheet.Cells[start, 1].Text.ToString() == "ПРИНАДЛЕЖНОСТИ") break;
                start++;
            }
            int fin = start;
            while (true)
            {
                if (ObjWorkSheet.Cells[fin, 1].Text.ToString() == "") break;
                fin++;
            }

            string frontParts = "";
            string backParts = "";
            string hand = "";
            int mas = 0;
            bool boolHand = false;
            int numberStr = 1;
            for (int i = start + 1; i < fin; i++)
            {
                string rez = ObjWorkSheet.Cells[i, 1].Text.ToString();
                while (ObjWorkSheet.Cells[i + 1, 1].Text.ToString().IndexOf((numberStr + 1).ToString() + ". ") != 0 && i < fin)
                {
                    i++;
                    rez += " " + ObjWorkSheet.Cells[i, 1].Text.ToString().Replace("\"", "");
                }
                rez = rez.Substring(numberStr.ToString().Length + 2);
                if (Array.IndexOf(numberPartsExcel, numberStr.ToString()) != -1)
                {
                    string numberPart = getFrontPart(rez);
                    if (numberPart != "0")
                    {
                        frontParts += numberPart + ";";
                        mas += 5;
                    }
                    numberPart = getBackPart(rez);
                    if (numberPart != "0")
                    {
                        backParts += numberPart + ";";
                        mas += 5;
                    }
                    if (rez == handString)
                    {
                        hand = "5";
                        boolHand = true;
                    }
                }
                numberStr++;
            }
            

            //todo remove
            GC.Collect();

            int[] values = { 20, 25, 40, 50, 80, 100, 150, 200 };
            string parClap = getData(Array.IndexOf(values, DN), form.ToLower(), asme);
            bool privPar = false;
            if (priv[0] == "88") privPar = true;
            ram_par[0] += (getWeight(Array.IndexOf(values, DN), asme) + getWeightPriv(Convert.ToInt32(priv[1]), boolHand) + mas).ToString();
            string ram_par_str = "";
            foreach (string i in ram_par)
            {
                ram_par_str += i + ";";
            }
            ram_par_str.Trim(';');
            string s = await RequestAsync(fileNameRezDvg, hand, parClap, ram_par_str, getDataPriv(Convert.ToInt32(priv[1]), privPar), backParts.Trim(';'), frontParts.Trim(';'));
            MessageBox.Show(s);
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
rez = reader.ReadToEnd();
response.Close();
return rez;
}*/
    }

}
