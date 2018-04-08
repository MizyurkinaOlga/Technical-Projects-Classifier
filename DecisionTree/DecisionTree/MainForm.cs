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


            //this.label1.MinimumSize = new Size(sizeFormX / 10, 0);
            //this.label1.MaximumSize = new Size(sizeFormX - (sizeFormX / 5), sizeFormY);
            //this.label1.Location = new Point((sizeFormX - this.label1.Size.Width) / 2, 30);

            //this.label3.MinimumSize = new Size(sizeFormX / 10, 0);
            //this.label3.MaximumSize = new Size(sizeFormX / 3 * 2, sizeFormY);
            //this.label3.Location = new Point((sizeFormX - this.label3.Size.Width) / 2, sizeFormY / 20 * 6);

            //this.button1.Location = new Point(sizeFormX - (sizeFormX / 4), (sizeFormY * 2) / 5);
            //this.label2.Text = "";
            //this.textBox1.Location = new Point(sizeFormX - (sizeFormX / 4) - this.textBox1.Size.Width - 20, (sizeFormY * 2) / 5);
            //this.label2.Location = new Point(sizeFormX - (sizeFormX / 4) - this.textBox1.Size.Width - 20, ((sizeFormY * 2) / 5) + 40);

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
                    //textBox1.Text = nameFile;
                    //label2.Text = "Загрузка файла. Подождите...";
                    //формируем объект для работы с файлом Excel
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
                            
                            //определяем в каком диапазоне документа записаны данные
                            int column = ExcSheet.Cells.SpecialCells(MExcel.XlCellType.xlCellTypeLastCell).Column;
                            int rows = ExcSheet.Cells.SpecialCells(MExcel.XlCellType.xlCellTypeLastCell).Row;
                            int firstColoumn = 1, firstRow = 1;
                            for (int i = 1; i < rows; i++)
                            {
                                bool fl = false;
                                for (int j = 1; j < column; j++)
                                {
                                    if (ExcSheet.Cells[i, j].Value != null)
                                    {
                                        firstColoumn = j;
                                        firstRow = i;
                                        fl = true;
                                        break;
                                    }
                                }
                                if (fl) break;
                            }
                            //формируем в памяти программы массивы и матрицы из исходных данных для дальнейшей работы
                            data = new string[rows - firstRow, column - firstColoumn];
                            inputs = new string[column - firstColoumn];
                            outputs = new string[rows - firstRow];
                            for (int j = 0; j < column - firstColoumn; j++)
                            {
                                inputs[j] = ExcSheet.Cells[firstRow, j + firstColoumn].Text;
                            }
                            for (int i = 0; i < rows - firstRow; i++)
                            {
                                for (int j = 0; j < column - firstColoumn; j++)
                                {
                                    data[i, j] = ExcSheet.Cells[i + firstRow + 1, j + firstColoumn].Text;
                                }
                            }
                            for (int i = 0; i < rows - firstRow; i++)
                            {
                                outputs[i] = ExcSheet.Cells[i + firstRow + 1, column].Text;
                            }
                            //label2.Text = "Чтение из файла завершено.";
                            //после чтения закрываем файл и завершаем работу с Excel
                            ExcelBook.Close();
                            ObjExcel.Quit();

                            //НАДО РИСОВАТЬ ГРАФИКИ
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
