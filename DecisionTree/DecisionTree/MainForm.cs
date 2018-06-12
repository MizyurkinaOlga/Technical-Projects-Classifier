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
        Dictionary<string, Dictionary<string, int>> allUniqInputs;
        Dictionary<string, Dictionary<string, Dictionary<double, double>>> justificationOfFuzzySet;
        Dictionary<string,Color> colorForNodes;
        Dictionary<string, Dictionary<string, double>> conformityStringToDouble;
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

                            MExcel.Worksheet ExcSheet = ExcelBook.Sheets[table.Selection];

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
                            justificationOfFuzzySet = new Dictionary<string, Dictionary<string, Dictionary<double, double>>>();
                            allUniqInputs = new Dictionary<string, Dictionary<string, int>>();
                            typeOfInputs = Utilities.TypeOfInputs(data, inputs);
                            conformityStringToDouble = new Dictionary<string, Dictionary<string, double>>();

                            comboBox1.Items.Clear();
                            comboBox1.Items.AddRange(inputs);
                            comboBox1.Enabled = true;

                            colorForNodes = GetColorForNodes(outputs);

                            ExcelBook.Close();
                            ObjExcel.Quit();
                            label1.Text = "";
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
            for (int i = 0; i < cls.Count; i++)
            {
                forret.Add(cls[i], colorList[i]);
            }
            //forret.Add(cls[0], Color.Blue);
            //forret.Add(cls[1], Color.Black);
            //forret.Add(cls[2], Color.Green);
            return forret;
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
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
                if (typeOfInputs[inputs[attributeIndex]] == "string")
                {
                    StringToDouble strToDouble = new StringToDouble(uniqValues.Keys.ToList());
                    if (strToDouble.ShowDialog(this) == DialogResult.OK)
                    {
                        conformityStringToDouble.Add(inputs[attributeIndex], strToDouble.Identify());
                    }
                }
                else
                {
                    conformityStringToDouble.Add(inputs[attributeIndex], Utilities.Confirm(uniqValues.Keys.ToList()));
                }
                if (comboBox2.SelectedItem.ToString()=="Прямой групповой")//прямой групповой метод
                {
                    //определение какой X к какому рангу -еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (method == 1)//статистических данных
                {
                    //определение какой X к какому рангу - еще одна форма и переписать UniqValCount
                    //формирование центров функции (только треугольные будем использовать)
                }
                if (comboBox2.SelectedItem.ToString() == "Равномерное покрытие")//равномерное покрытие
                {
                    centersFP = Utilities.CntrMFUniCover(ranks, conformityStringToDouble[inputs[attributeIndex]].Values.ToArray());
                    justificationOfFuzzySet.Add(inputs[attributeIndex],
                            Utilities.DegreeOfMembDouble(centersFP, conformityStringToDouble[inputs[attributeIndex]].Values.ToList()));
                    return centersFP;
                }
                if (comboBox2.SelectedItem.ToString() == "Случайное покрытие")//случайное покрытие
                {
                    //формирование центров функции (только треугольные будем использовать)
                    centersFP = Utilities.CntrMFRandomCover(ranks, conformityStringToDouble[inputs[attributeIndex]].Values.ToArray());
                    justificationOfFuzzySet.Add(inputs[attributeIndex],
                            Utilities.DegreeOfMembDouble(centersFP, conformityStringToDouble[inputs[attributeIndex]].Values.ToList()));
                    return centersFP;
                }
                if (comboBox2.SelectedItem.ToString() == "Частотный анализ значений")//для лингвистических переменных
                {
                    Dictionary<string, List<string>> valToRanks;
                    ExpertReview exprtRev = new ExpertReview(inputs[attributeIndex], ranks, uniqValues);
                    if (exprtRev.ShowDialog(this) == DialogResult.OK)
                    {
                        valToRanks = exprtRev.RetValToRanks();
                        centersFP = Utilities.CntrMFLingVar(valToRanks, uniqValues);
                        //Dictionary<string, Dictionary<string, double>> tmpDegre = Utilities.DegreeOfMembLing(centersFP, valToRanks, allUniqInputs[inputs[attributeIndex]]);
                        List<double> setZeroOne = Utilities.ReflectionOnTheZeroOne(conformityStringToDouble[inputs[attributeIndex]].Values.ToList());
                        Dictionary<string, Dictionary<double, double>> degr = Utilities.DegreeOfMembDouble(centersFP, setZeroOne);
                        Dictionary<string, Dictionary<double, double>> forJust = Utilities.ReturnToAB(degr, conformityStringToDouble[inputs[attributeIndex]].Values.ToList());
                        justificationOfFuzzySet.Add(inputs[attributeIndex], forJust);
                        int a = 5;
                    }
                    //
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
                if (typeOfInputs[inputs[j]] == "string" && i==0)
                {

                    list = ListPointONEF(znach, i);
                }
                else
                {
                    if (typeOfInputs[inputs[j]] == "string" && i == znach.Count - 1)
                    {
                        list = ListPointONEL(znach, i);
                    }
                    else
                    {
                        list = ListPointZERO(znach, i);
                    }                    
                }                
                
                LineItem graph = panel.AddCurve(allCentersOfMembFunc[inputs[j]].Keys.ToList()[i], list,
                    colorList[i], SymbolType.Star);
            }

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }
        private PointPairList ListPointZERO(List<List<double>> znach, int index)
        {
            PointPairList list = new PointPairList();
            list.Add(znach[index][0], 0.0);
            list.Add(znach[index][1], 1.0);
            list.Add(znach[index][2], 0.0);
            return list;
        }
        private PointPairList ListPointONEF(List<List<double>> znach, int index)
        {
            PointPairList list = new PointPairList();
            list.Add(znach[index][0], 1.0);
            list.Add(znach[index][1], 1.0);
            list.Add(znach[index][2], 0.0);
            return list;
        }
        private PointPairList ListPointONEL(List<List<double>> znach, int index)
        {
            PointPairList list = new PointPairList();
            list.Add(znach[index][0], 0.0);
            list.Add(znach[index][1], 1.0);
            list.Add(znach[index][2], 1.0);

            return list;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = false;
            comboBox2.Text = "Выберите метод...";
            comboBox2.Enabled = true;            
        }
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
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
            label4.Text = "";
            foreach(var item in colorForNodes)
            {
                label4.Text += item.Key + "->" + item.Value.ToString() + "; ";
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
                        double tmp = conformityStringToDouble[inputs[j]][data[i, j]];
                        double prob = justificationOfFuzzySet[inputs[j]][rank][tmp];
                        if (prob >= 0.5)
                        {
                            indexes.Add(i);
                        }
                    }
                    forRanks.Add(rank, indexes);
                }
            }
        }
        private void BuildSubTree(DecisionNode rt, Dictionary<string, Dictionary<string, Dictionary<double, double>>> fuzzySets)
        {
            List<DecisionNode> ListNodeXk = new List<DecisionNode>();
            List<double> GainRatio = new List<double>();
            List<double> Gain = new List<double>();
            List<double> SplitInfo = new List<double>();
            for (int i = 0; i < fuzzySets.Count; i++)
            {
                GainRatio.Add(0.0);
                Gain.Add(0.0);
                SplitInfo.Add(0.0);
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
                        double val = conformityStringToDouble[inputs[j]][data[i, j]];
                        if (rank.Value[val] >= 0.5)
                        {
                            tmp.ListIndexElements.Add(i);
                        }
                    }
                    tmp.ProbabilityClasses = Utilities.ProbabilityOfClass(new HashSet<string>(outputs).ToList(), tmp.ListIndexElements, outputs);
                    tmp.Entropy = Utilities.Info(tmp.ProbabilityClasses.Values.ToList());
                    Gain[k] += (Convert.ToDouble(tmp.ListIndexElements.Count) /
                                Convert.ToDouble(rt.ListIndexElements.Count))
                                        * tmp.Entropy;
                    double fraq = (double) tmp.ListIndexElements.Count / (double) rt.ListIndexElements.Count;
                    double lg = Math.Log(fraq) / Math.Log(2);
                    SplitInfo[k] -= fraq * (Math.Log(fraq) / Math.Log(2));
                    ListNodeXk.Add(tmp);
                }
                //вычислить энтропию переменной
                Gain[k] = rt.Entropy - Gain[k];
                GainRatio[k] = Gain[k] / SplitInfo[k];
                k++;
            }
            //проверить что хотя бы в N дочерних больше чем E элементов
            int N = 2;
            int E = 2;
            int elemMoreThenE = 0;
            foreach(var item in ListNodeXk)
            {
                if (item.ListIndexElements.Count > E)
                {
                    elemMoreThenE++;
                }
            }
            if (elemMoreThenE > N)
            {
                int Xmax = inputs.ToList().IndexOf(fuzzySets.Keys.ToList()[GainRatio.IndexOf(GainRatio.Max())]);
                foreach (var nd in ListNodeXk)
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
                Dictionary<string, Dictionary<string, Dictionary<double, double>>> newFuzzySet = new Dictionary<string, Dictionary<string, Dictionary<double, double>>>();
                foreach (var item in fuzzySets)
                {
                    if (item.Key != inputs[Xmax])
                    {
                        newFuzzySet[item.Key] = item.Value;
                    }
                }
                if (newFuzzySet.Count > 0)
                {
                    foreach (DecisionNode curNod in rt.Nodes)
                    {
                        if (!curNod.ProbabilityClasses.Values.ToList().Contains(1.0))
                        {
                            BuildSubTree(curNod, newFuzzySet);
                        }
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
