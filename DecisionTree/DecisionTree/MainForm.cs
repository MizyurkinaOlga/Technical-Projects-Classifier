using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using Accord.IO;
using ZedGraph;

namespace DecisionTree
{
    public partial class MainForm : Form
    {
        public string[,] data;
        public string[] inputs;
        public string[] outputs;
        string pathToFileMembFunc;
        Dictionary<string, string> typeOfInputs;

        Dictionary<string, Dictionary<string, double>> allCentersOfMembFunc;
        public MainForm()
        {
            InitializeComponent();
            data = new string[10, 10];
            inputs = new string[1];
            outputs = new string[1];
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            int sizeFormX = 800;
            int sizeFormY = 500;
            this.Size = new Size(sizeFormX, sizeFormY);

        }

        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Excel File";
            openFileDialog1.Filter = "Excel Worksheets (*.xls;*.xlsx)|*.xls;*.xlsx|All File (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var nameFile = openFileDialog1.FileName;

                    //form an object to work with an Excel file
                    string extension = Path.GetExtension(nameFile);
                    if (extension == ".xls" || extension == ".xlsx")
                    {
                        MExcel.Application ObjExcel = new MExcel.Application();
                        MExcel.Workbook ExcelBook = ObjExcel.Workbooks.Open(nameFile);

                        List<string> worksheets = new List<string>();
                        foreach(MExcel.Worksheet each in ExcelBook.Worksheets)
                        {
                            worksheets.Add(each.Name);
                        }
                        SelectTableSheet table = new SelectTableSheet(worksheets.ToArray());
                        if (table.ShowDialog(this) == DialogResult.OK)
                        {
                            MExcel.Worksheet ExcSheet = ExcelBook.Sheets[worksheets.Count() - table.Selection];

                            //determining the range of data storage in a file
                            Dictionary<int, int> RC = Utilities.RangeOfData(ExcSheet);
                            int rows = RC.First().Key;
                            int column = RC.First().Value;
                            int firstRow = RC.Last().Key;
                            int firstColumn = RC.Last().Value;

                            //recording in the program memory arrays and matrix from the original data for further work
                            data = new string[rows, column];
                            inputs = new string[column];
                            outputs = new string[rows];
                            for (int j = 0; j < column; j++)
                            {
                                inputs[j] = ExcSheet.Cells[firstRow, j + firstColumn].Text;
                            }
                            comboBox1.Items.AddRange(inputs);
                            for (int i = 0; i < rows; i++)
                            {
                                for (int j = 0; j < column; j++)
                                {
                                    data[i, j] = ExcSheet.Cells[i + firstRow + 1, j + firstColumn].Text;
                                }
                            }
                            for (int i = 0; i < rows; i++)
                            {
                                outputs[i] = ExcSheet.Cells[i + firstRow + 1, column + firstColumn].Text;
                            }
                            pathToFileMembFunc = Environment.CurrentDirectory +
                                        "\\MembershipFunction\\" +
                                        ExcSheet.Name + ".txt";
                            allCentersOfMembFunc = new Dictionary<string, Dictionary<string, double>>();
                            typeOfInputs = Utilities.TypeOfInputs(data, inputs);
                            comboBox1.Enabled = true;

                            ExcelBook.Close();
                            ObjExcel.Quit();
                        }
                    }
                }
                catch (Exception excp)
                {
                    MessageBox.Show(excp.Message.ToString());
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private Dictionary<string, double> DefineMethods (int attributeIndex, string typeInput)
        {
            int column = inputs.Length;
            int rows = outputs.Length;
            string[] attributeValues = new string[rows];
            for (int i = 0; i < rows; i++)
            {
                attributeValues[i] = data[i, attributeIndex];
            }
            Dictionary<string, double> centersFP = new Dictionary<string, double>();
            Dictionary<string, int> uniqValues = Utilities.UniqValCount(attributeValues);//!!!!!
            DefineRanks ranksForm = new DefineRanks(inputs[attributeIndex]);
            if (ranksForm.ShowDialog(this) == DialogResult.OK)
            {
                List<string> ranks = ranksForm.Identify();

                int method = comboBox2.SelectedIndex;
                
                if (method == 0)//прямой групповой метод
                {
                    //определение какой X к какому рангу -еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 1)//статистических данных
                {
                    //определение какой X к какому рангу - еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 2)//равномерное покрытие
                {
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 3)//случайное покрытие
                {
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 4)//для лингвистических переменных
                {
                    //определение какой X к какому рангу - еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }
                



                //List<string> orderValues = ranks.OrderedValues();
                //List<int> tmpOrderCount = new List<int>();
                //foreach (var item in orderValues)
                //{
                //    tmpOrderCount.Add(uniqValues[item]);
                //}
                //uniqValues.Clear();
                //for (int k = 0; k < orderValues.Count; k++)
                //{
                //    uniqValues.Add(orderValues[k], tmpOrderCount[k]);
                //}
            }
            centersFP = Utilities.CntrMFLingVar(uniqValues, attributeValues.Length);
            return centersFP;
        }
        private void WriteCentersToFile()
        {
            string inFile = "";
            foreach (var attr in allCentersOfMembFunc)
            {
                inFile += attr.Key + " = { " + "\r\n";
                foreach (var item in attr.Value)
                {
                    inFile += '\u0022' + item.Key + '\u0022' + " : " + '\u0022' + item.Value.ToString() + '\u0022' + "\r\n";
                }
                inFile += "};\r\n\r\n";
            }
            Byte[] info = new UTF8Encoding(true).GetBytes(inFile);
            FileStream fileMembFunc = File.Create(pathToFileMembFunc);
            fileMembFunc.Write(info, 0, info.Length);
            fileMembFunc.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int j = comboBox1.SelectedIndex;
            if (allCentersOfMembFunc.ContainsKey(inputs[j]))
            {
                OverwritingFunction overwrite = new OverwritingFunction(inputs[j]);
                if (overwrite.ShowDialog(this) == DialogResult.OK)
                {
                    allCentersOfMembFunc[inputs[j]] = DefineMethods(j, typeOfInputs[inputs[j]]);
                }
            }
            else
            {
                allCentersOfMembFunc.Add(inputs[j], DefineMethods(j, typeOfInputs[inputs[j]]));
            }

            WriteCentersToFile();


            GraphPane panel = zedGraphControl1.GraphPane;
            panel.Title.Text = inputs[j];
            panel.XAxis.Title.Text = "Значение аттрибута";
            panel.YAxis.Title.Text = "Значение ФП";
            panel.CurveList.Clear();


            List<double> znach = allCentersOfMembFunc[inputs[j]].Values.ToList();
            HashSet<Color> colorList = new HashSet<Color>();
            Random color = new Random();
            int h = 0;
            while (h < znach.Count)
            {
                int r = color.Next() % 2;
                int g = color.Next() % 2;
                int b = color.Next() % 2;
                if (r==b && b==g && g == 1)
                {
                    r = 0;
                }
                if (colorList.Add(Color.FromArgb(r * 255, g * 255, b * 255)))
                {
                    h++;
                }
            }
            for(int i = 0; i < znach.Count; i++)
            {
                PointPairList list = new PointPairList();
                list.Add(-0.1, 0.0);
                if (i == 0)
                {
                    list.Add(0.0, 0.0);
                }
                else
                {
                    list.Add(znach[i - 1], 0.0);
                }

                list.Add(znach[i], 1.0);
                if (i < znach.Count - 1)
                {
                    list.Add(znach[i + 1], 0.0);
                }
                else
                {
                    list.Add(1.0, 0.0);
                }
                list.Add(1.1, 0.0);
                LineItem graph = panel.AddCurve(allCentersOfMembFunc[inputs[j]].Keys.ToList()[i], list,
                    colorList.ToList()[i], SymbolType.Star);
            }

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();

            //НАДО РИСОВАТЬ ГРАФИКИ
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (typeOfInputs[inputs[comboBox1.SelectedIndex]] == "string")
            {
                comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            }
            else
            {
                comboBox2.Enabled = true;
            }
        }
    } 
}
