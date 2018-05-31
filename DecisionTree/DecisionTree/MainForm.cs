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
        Dictionary<string, Dictionary<string, List<double>>> allCentersOfMembFunc;
        Dictionary<string, Dictionary<string, int>> allUniqInputs;//атрибут,значение,количество таких
        Dictionary<string, Dictionary<string, Dictionary<string, double>>> justificationOfFuzzySet;
        //поиск по атрибут->ранг->значение атрибута->степень принадлежности
        Dictionary<string, Dictionary<string, List<int>>> fuzzySets;
        //атрибут->ранг->список индексов из таблицы data
        Dictionary<string,Color> colorForNodes;
        int pictureIndex = 0;
        double limitDecision = 0.98;
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
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox1.Text = "Выберите атрибут...";
                    comboBox2.Text = "Выберите метод...";
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
                            pathToFileMembFunc = Environment.CurrentDirectory +
                                        "\\MembershipFunction\\" +
                                        ExcSheet.Name + ".txt";
                            //
                            allCentersOfMembFunc = new Dictionary<string, Dictionary<string, List<double>>>();
                            justificationOfFuzzySet = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                            allUniqInputs = new Dictionary<string, Dictionary<string, int>>();
                            typeOfInputs = Utilities.TypeOfInputs(data, inputs);
                            
                            comboBox1.Items.Clear();
                            comboBox1.Items.AddRange(inputs);
                            comboBox1.Enabled = true;

                            colorForNodes = GetColorForNodes(outputs);

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
        private Dictionary<string,Color> GetColorForNodes(string[] outputs)
        {
            HashSet<string> cls1 = new HashSet<string>(outputs);
            List<string> cls = cls1.ToList();
            List<Color> colorList = GetColor(cls.Count);
            Dictionary<string, Color> forret = new Dictionary<string, Color>();
            for(int i = 0; i < cls.Count; i++)
            {
                forret.Add(cls[i], colorList[i]);
            }
            return forret;
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private Dictionary<string,Dictionary<string,double>> DegreeOfMembDouble(Dictionary<string, List<double>> centers, int index)
        {
            Dictionary<string, Dictionary<string, double>> rankDegreeze = new Dictionary<string, Dictionary<string, double>>();
            foreach (var rank in centers)
            {
                Dictionary<string, double> degreeze = new Dictionary<string, double>();
                foreach (var uniqZn in allUniqInputs[inputs[index]])
                {
                    if (Convert.ToDouble(uniqZn.Key) < rank.Value[0] || Convert.ToDouble(uniqZn.Key) > rank.Value[2])
                    {
                        degreeze.Add(uniqZn.Key, 0.00);
                    }
                    else
                    {
                        if (Convert.ToDouble(uniqZn.Key) == rank.Value[1])//если равно центру
                        {
                            degreeze.Add(uniqZn.Key, 1.00);
                        }
                        else
                        {
                            if (Convert.ToDouble(uniqZn.Key) < rank.Value[1])
                            {
                                degreeze.Add(uniqZn.Key,
                                    (Convert.ToDouble(uniqZn.Key) - rank.Value[0]) / (rank.Value[1] - rank.Value[0]));
                            }
                            else
                            {
                                degreeze.Add(uniqZn.Key,
                                    (rank.Value[2] - Convert.ToDouble(uniqZn.Key)) / (rank.Value[2] - rank.Value[1]));
                            }
                        }
                    }
                }
                rankDegreeze.Add(rank.Key, degreeze);
            }
            return rankDegreeze;
        }
        private Dictionary<string, List<double>> DefineMethods (int attributeIndex, string typeInput)
        {
            int column = inputs.Length;
            int rows = outputs.Length;
            string[] attributeValues = new string[rows];
            for (int i = 0; i < rows; i++)
            {
                attributeValues[i] = data[i, attributeIndex];
            }
            Dictionary<string, List<double>> centersFP = new Dictionary<string, List<double>>();
            Dictionary<string, int> uniqValues = Utilities.UniqValCount(attributeValues);
            allUniqInputs.Add(inputs[attributeIndex], uniqValues);
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
                    centersFP = Utilities.CntrMFUniCover(ranks, attributeValues);
                    justificationOfFuzzySet.Add(inputs[attributeIndex], DegreeOfMembDouble(centersFP, attributeIndex));
                    return centersFP;
                }
                if (method == 3)//случайное покрытие
                {
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 4)//для лингвистических переменных
                {
                    //return Utilities.CntrMFLingVar(ranks, attributeValues);
                    //определение какой X к какому рангу - еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }                
            }
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
                    inFile += '\u0022' + item.Key + '\u0022' + " : " + '\u0022' + item.Value[1].ToString() + '\u0022' + "\r\n";
                }
                inFile += "};\r\n\r\n";
            }
            Byte[] info = new UTF8Encoding(true).GetBytes(inFile);
            FileStream fileMembFunc = File.Create(pathToFileMembFunc);
            fileMembFunc.Write(info, 0, info.Length);
            fileMembFunc.Close();
        }
        private List<Color> GetColor(int countColor)
        {
            HashSet<Color> colorList = new HashSet<Color>();
            Random color = new Random();
            int h = 0;
            while (h < countColor)
            {
                int r = color.Next() % 2;
                int g = color.Next() % 2;
                int b = color.Next() % 2;
                if (r == b && b == g && g == 1)
                {
                    r = 0;
                }
                if (colorList.Add(Color.FromArgb(r * 255, g * 255, b * 255)))
                {
                    h++;
                }
            }
            return colorList.ToList();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int j = inputs.ToList().FindIndex(x => x == comboBox1.SelectedItem.ToString());
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


            List<List<double>> znach = allCentersOfMembFunc[inputs[j]].Values.ToList();
            List<Color> colorList = GetColor(znach.Count);
            for(int i = 0; i < znach.Count; i++)
            {
                PointPairList list = new PointPairList();
                list.Add(znach[i][0] - 0.1, 0.0);
                list.Add(znach[i][0], 0.0);
                list.Add(znach[i][1], 1.0);
                list.Add(znach[i][2], 0.0);
                list.Add(znach[i][2] + 0.1, 0.0);
                
                LineItem graph = panel.AddCurve(allCentersOfMembFunc[inputs[j]].Keys.ToList()[i], list,
                    colorList.ToList()[i], SymbolType.Star);
            }

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = false;
            comboBox2.Text = "Выберите метод...";
            if (typeOfInputs[comboBox1.SelectedItem.ToString()] == "string")
            {
                comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            }
            else
            {
                comboBox2.Enabled = true;
            }
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            fuzzySets = new Dictionary<string, Dictionary<string, List<int>>>();        
            label3.Text = "функции принадлежности построены для " + allCentersOfMembFunc.Count() + " атрибутов";
            if (allCentersOfMembFunc.Count() < inputs.Length)
            {
                for (int j = 0; j < inputs.Length; j++)
                {
                    if (!allCentersOfMembFunc.ContainsKey(inputs[j]))
                    {
                        allCentersOfMembFunc.Add(inputs[j], DefineMethods(j, typeOfInputs[inputs[j]]));
                    }
                    label3.Text = "функции принадлежности построены для " + allCentersOfMembFunc.Count() + " атрибутов";
                }
                WriteCentersToFile();
            }
            for(int j = 0; j < inputs.Length; j++)
            {
                Dictionary<string, List<int>> forRanks = new Dictionary<string, List<int>>();
                foreach(var rank in justificationOfFuzzySet[inputs[j]].Keys)
                {
                    List<int> indexes = new List<int>();
                    int p = data.Length / inputs.Length;
                    for (int i = 0; i < p; i++)
                    {
                        if (justificationOfFuzzySet[inputs[j]][rank][data[i, j]] >= 0.5)
                        {
                            indexes.Add(i);
                        }
                    }
                    forRanks.Add(rank, indexes);
                }
                fuzzySets.Add(inputs[j], forRanks);
            }
        }
        private void BuildSubTree(DecisionNode rt, Dictionary<string, Dictionary<string, Dictionary<string, double>>> fuzzySets)
        {
            List<DecisionNode> ListNodeXk = new List<DecisionNode>();
            List<double> GainRatio = new List<double>();
            for (int i = 0; i < fuzzySets.Count; i++)
            {
                GainRatio.Add(0.0);
            }
            int k = 0;
            foreach(var atribut in fuzzySets)
            {
                int j = Array.FindIndex(inputs, s => s.Equals(atribut.Key));
                
                foreach (var rank in atribut.Value)
                {
                    
                    DecisionNode tmp = new DecisionNode(atribut.Key, rank.Key, new List<int>(), 0.0, -1, new Dictionary<string, double>());
                    foreach(var i in rt.ListIndexElements)
                    {
                        if (rank.Value[data[i, j]] >= 0.5)
                        {
                            tmp.ListIndexElements.Add(i);
                        }
                    }
                    tmp.ProbabilityClasses = Utilities.ProbabilityOfClass(new HashSet<string>(outputs).ToList(), tmp.ListIndexElements, outputs);
                    tmp.Entropy = Utilities.Info(tmp.ProbabilityClasses.Values.ToList());
                    GainRatio[k] += (Convert.ToDouble(tmp.ListIndexElements.Count) /
                                    Convert.ToDouble(rt.ListIndexElements.Count))
                                        * tmp.Entropy;
                    ListNodeXk.Add(tmp);
                }
                //вычислить энтропию переменной
                k++;
            }
            int Xmax = inputs.ToList().IndexOf(fuzzySets.Keys.ToList()[GainRatio.IndexOf(GainRatio.Min())]);
            foreach(var nd in ListNodeXk)
            {
                if (nd.Atribute == inputs[Xmax])
                {
                    Bitmap pic = Utilities.CreatePicturePercent(colorForNodes, nd.ProbabilityClasses);
                    treeView1.ImageList.Images.Add(pic);
                    nd.ImageIndex = pictureIndex;
                    pictureIndex++;
                    rt.Nodes.Add(nd);
                }
            }
            Dictionary<string, Dictionary<string, Dictionary<string, double>>> newFuzzySet = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
            foreach(var item in fuzzySets)
            {
                if (item.Key!= inputs[Xmax])
                {
                    newFuzzySet[item.Key]=item.Value;
                }
            }
            if (newFuzzySet.Count > 0)
            {
                foreach(DecisionNode curNod in rt.Nodes)
                {
                    if (!curNod.ProbabilityClasses.Values.ToList().Contains(1.0))
                    {
                        BuildSubTree(curNod, newFuzzySet);
                    }
                }
            }
        }
        private void btnBuildTree_Click(object sender, EventArgs e)
        {
            List<int> listIndexSet = new List<int>();
            for (int i = 0; i < outputs.Length; i++)
            {
                listIndexSet.Add(i);
            }
            Dictionary<string, double> probability = Utilities.ProbabilityOfClass(new HashSet<string>(outputs).ToList(), listIndexSet, outputs);
            double infoSet = Utilities.Info(probability.Values.ToList());

            Bitmap pic = Utilities.CreatePicturePercent(colorForNodes, probability);
            treeView1.ImageList = new ImageList();
            treeView1.ImageList.Images.Add(pic);

            DecisionNode root = new DecisionNode("ALL", "", listIndexSet, infoSet, pictureIndex, probability);
            pictureIndex++;
            BuildSubTree(root, justificationOfFuzzySet);

            treeView1.Nodes.Add(root);            
            treeView1.Update();
        }
    } 
}
