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

namespace DecisionTree
{
    public partial class MainForm : Form
    {
        public string[,] data;
        public string[] inputs;
        public string[] outputs;
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
                            string path = Environment.CurrentDirectory + 
                                        "\\MembershipFunction\\" + 
                                        ExcSheet.Name + ".txt";
                            FileStream memberFunct = File.Create(path);
                            string inFile = "";

                            ExcelBook.Close();
                            ObjExcel.Quit();

                            //difine type of inputs
                            Dictionary<string, string> typeOfInputs = Utilities.TypeOfInputs(data, inputs);                            

                            for (int j = 0; j < column; j++)
                            {
                                string[] attributeValues = new string[rows];
                                for (int i = 0; i < rows; i++)
                                {
                                    attributeValues[i] = data[i, j];
                                }
                                Dictionary<string, double> centersFP = new Dictionary<string, double>();
                                if (typeOfInputs[inputs[j]] == "string")
                                {
                                    Dictionary<string, int> uniqValues = Utilities.UniqValCount(attributeValues);
                                    SortRanks ranks = new SortRanks(inputs[j], uniqValues.Keys.ToArray());
                                    if (ranks.ShowDialog(this) == DialogResult.OK)
                                    {
                                        List<string> orderValues = ranks.OrderedValues();
                                        List<int> tmpOrderCount = new List<int>();
                                        foreach (var item in orderValues)
                                        {
                                            tmpOrderCount.Add(uniqValues[item]);
                                        }
                                        uniqValues.Clear();
                                        for(int k = 0; k < orderValues.Count; k++)
                                        {
                                            uniqValues.Add(orderValues[k], tmpOrderCount[k]);
                                        }
                                    }
                                    centersFP = Utilities.CentersOfFP(uniqValues, attributeValues.Length);
                                }
                                else
                                {

                                }
                                inFile += inputs[j] + " = { " + "\r\n";
                                foreach(var item in centersFP)
                                {
                                    inFile += '\u0022' + item.Key + '\u0022' + " : "+ '\u0022' + item.Value.ToString() + '\u0022' + "\r\n";
                                }
                                inFile += "};\r\n";
                                
                            }

                            //НАДО РИСОВАТЬ ГРАФИКИ
                            Byte[] info = new UTF8Encoding(true).GetBytes(inFile);
                            memberFunct.Write(info, 0, info.Length);
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
    } 
}
