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
        List<String> excelPar = new List<String>();
        String filename;

        // Возмыжные варианты размеров DN.
        int[] values = { 20, 25, 40, 50, 80, 100, 150, 200 };
        public Form1()
        {
            InitializeComponent();
        }

        private async Task<String> RequestAsync(String filename, String hand, String valveParameters, String frameParameters, String actuator_par, String backParts, String frontParts)
        {
            StreamReader reader = null;
            try
            {
                String url = $"https://localhost:44394/Index?klap_par={valveParameters}&klap=Клапан&frameParameters={frameParameters}&actuator_par={actuator_par}&"
                    + $"backParts={backParts}&frontParts={frontParts}&hand={hand}&filename={filename}";
                WebRequest request = WebRequest.Create(url);
                WebResponse response = await request.GetResponseAsync().ConfigureAwait(true);
                Stream stream = response.GetResponseStream();
                reader = new StreamReader(stream);
                return await reader.ReadToEndAsync();
            } lineNumberAccessoriesally
            {
                reader.Close();
            }
        }


        public static String[][] getDataFromCSV(String pathCsvFile)
        {
            List<String[]> data = new List<String[]>();
            using (System.IO.StreamReader file = new System.IO.StreamReader(pathCsvFile))
            {
                String line;
                while ((line = file.ReadLine()) != null)
                {
                    String[] parts_of_line = line.Split(';');
                    String[] mass = new String[parts_of_line.Length];
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
        public static String getData(int value, String form, int class_asme)
        {
            String[][] dataMassCsvFileA = getDataFromCSV(@"..\..\..\..\csvFiles\A.csv");
            String[][] dataMassCsvFileB = getDataFromCSV(@"..\..\..\..\csvFiles\B.csv");
            String[][] dataMassCsvFileC = getDataFromCSV(@"..\..\..\..\csvFiles\C.csv");


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
        public static String getDataPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            String[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\Priv.csv");
            int parSize = 2;
            if (par) parSize++;
            return $"{dataMassCsvFile[1 + nomberSize][1]};" +
                    $"{dataMassCsvFile[1 + nomberSize][parSize]};" +
                    $"{dataMassCsvFile[1 + nomberSize][4]};" +
                    $"{dataMassCsvFile[1 + nomberSize][5]}";

        }
        public static int getWeightPriv(int size, bool par)
        {
            int[] sizes = new int[] { 6, 10, 16, 23 };
            int nomberSize = Array.IndexOf(sizes, size);
            String[][] dataMassCsvFile = getDataFromCSV(@"..\..\..\..\csvFiles\weightsDimen.csv");
            int parSize = par ? 2 : 1;
            return Convert.ToInt32(dataMassCsvFile[nomberSize + 1][parSize]);

        }
        public static int getWeight(int value, int class_asme)
        {
            String[][] dataMassCsvFileA = getDataFromCSV(@"..\..\..\..\csvFiles\bredWeights.csv");
            for (int i = 0; i < dataMassCsvFileA[0].GetLength(0); i++)
            {
                if (dataMassCsvFileA[0][i] == class_asme.ToString())
                {
                    return (Int32.Parse(dataMassCsvFileA[value + 1][i]));
                }
            }
            return 0;

        }
        public static String getFrontPart(String Part)
        {
            Part = Part.Replace("\"", "").Trim().ToLower();
            String[][] dataMassCsvFileParts = getDataFromCSV(@"..\..\..\..\csvFiles\artFrontParts.csv");
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
        public static String getBackPart(String Part)
        {
            Part = Part.Replace("\"", "").Trim().ToLower();
            String[][] dataMassCsvFileParts = getDataFromCSV(@"..\..\..\..\csvFiles\artBackParts.csv");
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
            int lineNumberAccessories = start;
            while (true)
            {
                if (ObjWorkSheet.Cells[lineNumberAccessories, 2].Text.ToString() == "") break;
                listBox1.Items.Add(ObjWorkSheet.Cells[lineNumberAccessories, 2].Text.ToString());
                excelPar.Add("Модель клапана: " + ObjWorkSheet.Cells[lineNumberAccessories, 4].Text.ToString() +
                            "\nМодель привода: " + ObjWorkSheet.Cells[lineNumberAccessories, 18].Text.ToString() +
                            "\nКласс давления: " + ObjWorkSheet.Cells[lineNumberAccessories, 8].Text.ToString() + "PN , " +
                            ObjWorkSheet.Cells[lineNumberAccessories, 9].Text.ToString() + " ASME" +
                            "\nDN: " + ObjWorkSheet.Cells[lineNumberAccessories, 11].Text.ToString());
                lineNumberAccessories++;
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
        public async void create_Arm(Excel.Worksheet ObjWorkSheet, int numberString, String fileNameRezDvg)
        {
            // создаём строку для 
            String handString = "Ручной дублер";
            // массив для хранения параметров рамки 
            String[] frameParameters = {    "Масса, кг: ", "Заказчик: ", "Потребитель: ", "Установка: ", "Позиция: ",
                                    "Модель привода: ", "Модель клапана: " , "Размер: DN " , "присоединение: PN " , "ХЗ1",
                                    "ХЗ2", "ХЗ3" , "ХЗ4" , "ХЗ5" , "ХЗ6"   };

            // считываем класс давления
            int asme = Convert.ToInt32(ObjWorkSheet.Cells[numberString, 9].Text.ToString());

            // считываем диаметр трубы
            int DN = Convert.ToInt32(ObjWorkSheet.Cells[numberString, 11].Text.ToString());

            // форма присоединеия ( фланец или под приварку, если я правильно понял).
            String form = ObjWorkSheet.Cells[numberString, 10].Text.ToString();

            // считываем модель привода.
            String[] actuator = ObjWorkSheet.Cells[numberString, 18].Text.ToString().Split('/');

            // добавляем к параметрам их значение.
            frameParameters[1] += ObjWorkSheet.Cells[2, 3].Text.ToString();
            frameParameters[2] += ObjWorkSheet.Cells[3, 3].Text.ToString();
            frameParameters[3] += ObjWorkSheet.Cells[5, 3].Text.ToString();
            frameParameters[4] += ObjWorkSheet.Cells[numberString, 2].Text.ToString();
            frameParameters[5] += ObjWorkSheet.Cells[numberString, 18].Text.ToString();
            frameParameters[6] += ObjWorkSheet.Cells[numberString, 4].Text.ToString();
            frameParameters[7] += ObjWorkSheet.Cells[numberString, 11].Text.ToString();
            frameParameters[8] += ObjWorkSheet.Cells[numberString, 8].Text.ToString();
            String[] numberPartsExcel = ObjWorkSheet.Cells[numberString, 28].Text.ToString().Split(',');

            // тут мы ищем строку после которой перечисленны номаера принадлежностей и их названия.
            int start = 1;
            while (true)
            {
                if (ObjWorkSheet.Cells[start, 1].Text.ToString() == "ПРИНАДЛЕЖНОСТИ") break;
                start++;
            }
            int lineNumberAccessories = start;
            while (true)
            {
                if (ObjWorkSheet.Cells[lineNumberAccessories, 1].Text.ToString() == "") break;
                lineNumberAccessories++;
            }

            // созаём пустые строки для частей и ручного привода.
            String frontParts = "";
            String backParts = "";
            String hand = "";

            // переменная для расчёта массы.
            int mas = 0;
            bool boolHand = false;
            int numberStr = 1;
            for (int i = start + 1; i < lineNumberAccessories; i++)
            {
                String result = ObjWorkSheet.Cells[i, 1].Text.ToString();
                while (ObjWorkSheet.Cells[i + 1, 1].Text.ToString().IndexOf((numberStr + 1).ToString() + ". ") != 0 && i < lineNumberAccessories)
                {
                    i++;
                    result += " " + ObjWorkSheet.Cells[i, 1].Text.ToString().Replace('\"', null);
                }
                result = result.Substring(numberStr.ToString().Length + 2);
                // тут распределение частей на те которые нужно вставлять перед приводои и те которые после, его нужно переделать
                if (Array.IndexOf(numberPartsExcel, numberStr.ToString()) != -1)
                {
                    String numberPart = getFrontPart(result);
                    if (numberPart != "0")
                    {
                        frontParts += numberPart + ";";
                        mas += 5;
                    }
                    numberPart = getBackPart(result);
                    if (numberPart != "0")
                    {
                        backParts += numberPart + ";";
                        mas += 5;
                    }
                    if (result == handString)
                    {
                        hand = "5";
                        boolHand = true;
                    }
                }
                numberStr++;
            }
            

            // получаем размеры клапана
            String valveParameters = getData(Array.IndexOf(values, DN), form.ToLower(), asme);
            // вот тут я так и не вспомни почему 88, но скорее всего потому что только с приводами этой линейки мы работали
            bool actuatorParameters = false;
            if (actuator[0] == "88") actuatorParameters = true;
            // считаем общию массу
            frameParameters[0] += (getWeight(Array.IndexOf(values, DN), asme) + getWeightPriv(Convert.ToInt32(actuator[1]), boolHand) + mas).ToString();
            // переделываем массив параметров рамки в строку 
            String frameParametersString = "";
            foreach (String i in frameParameters)
            {
                frameParametersString += i + ";";
            }
            // убираем последнию точку с запятой из этой строки
            frameParametersString.Trim(';');
            String requestResult = await RequestAsync(fileNameRezDvg, hand, valveParameters, frameParametersString, getDataPriv(Convert.ToInt32(actuator[1]), actuatorParameters), backParts.Trim(';'), frontParts.Trim(';'));
            MessageBox.Show(requestResult);
        }
    }

}
