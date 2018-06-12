using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MExcel = Microsoft.Office.Interop.Excel;
using ZedGraph;
using System.Drawing;

namespace DecisionTree
{
    public static class Utilities
    {
        public static Dictionary<string, string> TypeOfInputs(string[,] data, string[] inputs)
        {
            Dictionary<string, string> typeOfInputs = new Dictionary<string, string>();
            for (int j = 0; j < inputs.Length; j++)
            {
                double tmp;
                if (data[0, j].Any(c => char.IsLetter(c)))
                {
                    typeOfInputs.Add(inputs[j], "string");
                }
                else
                {
                    if (data[0, j].Count(c => c == '.') > 1 || data[0, j].Count(c => c == ',') > 1)
                    {
                        typeOfInputs.Add(inputs[j], "fuzzySet");
                        if (data[0, j].Contains("."))
                        {
                            int maxI = data.Length / inputs.Length;
                            for (int i = 0; i < maxI; i++)
                            {
                                data[i, j] = data[i, j].Replace(',', ';');
                                data[i, j] = data[i, j].Replace('.', ',');
                            }
                        }
                    }
                    else
                    {
                        typeOfInputs.Add(inputs[j], "double");
                        if (data[0, j].Contains("."))
                        {
                            int maxI = data.Length / inputs.Length;
                            for (int i = 0; i < maxI; i++)
                            {
                                data[i, j] = data[i, j].Replace('.', ',');
                            }
                        }
                    }
                }
            }
            return typeOfInputs;
        }
        public static Dictionary<int, int> RangeOfData(MExcel.Worksheet ExcSheet)
        {
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
            rows = rows - firstRow;
            column = column - firstColoumn;
            Dictionary<int, int> ret = new Dictionary<int, int>();
            ret.Add(rows, column);
            ret.Add(firstRow, firstColoumn);
            return ret;
        }
        public static Dictionary<string, int> UniqValCount(string[] values)
        {
            int countVal = values.Length;
            SortedDictionary<string, int> sortedUniq = new SortedDictionary<string, int>();
            foreach (var val in values)
            {
                if (sortedUniq.ContainsKey(val))
                {
                    sortedUniq[val] += 1;
                }
                else
                {
                    sortedUniq.Add(val, 1);
                }
            }
            Dictionary<string, int> uniqVal = new Dictionary<string, int>(sortedUniq);

            return uniqVal;
        }        
        public static Dictionary<string, List<double>> CntrMFUniCover(List<string> ranks, double[] values)
        {
            Dictionary<string, List<double>> centers = new Dictionary<string, List<double>>();

            List<double> valDouble = new List<double>();
            foreach (double item in values)
            {
                valDouble.Add(item);
            }
            double min = valDouble.Min();
            double max = valDouble.Max();
            double delta = (max - min) / (ranks.Count - 1);
            for (int j = 1; j <= ranks.Count; j++)
            {
                List<double> cntr = new List<double>();
                if (j == 1)
                {
                    cntr.Add(min);
                }
                else
                {
                    cntr.Add(min + (j - 2) * delta);
                }
                cntr.Add(min + (j - 1) * delta);
                if (j == ranks.Count)
                {
                    cntr.Add(max);
                }
                else
                {
                    cntr.Add(min + (j) * delta);
                }
                centers.Add(ranks[j - 1], cntr);
            }
            return centers;
        }
        public static Dictionary<string, List<double>> CntrMFRandomCover(List<string> ranks, double[] values)
        {
            Dictionary<string, List<double>> centers = new Dictionary<string, List<double>>();
            SortedSet<double> points = new SortedSet<double>();
            int count = ranks.Count * 3 - 2;
            Random val = new Random();
            double min = values.Min();
            double max = values.Max();
            while (points.Count < count)
            {
                double tmp = val.NextDouble();
                points.Add(min + (max - min) * tmp);
            }
            double[,] tmpCenters = new double[ranks.Count, 3];
            tmpCenters[0, 0] = min;
            tmpCenters[ranks.Count - 1, 2] = max;
            for (int i=0; i < count; i++)
            {
                int ost = i % 3;
                int cel = (int)i / 3;
                if (ost == 1)
                {
                    tmpCenters[cel + 1, 0] = points.ElementAt(i);
                }
                else
                {
                    if (ost == 0)
                    {
                        tmpCenters[cel, 1] = points.ElementAt(i);
                    }
                    else
                    {
                        tmpCenters[cel, 2] = points.ElementAt(i);
                    }
                }
            }
            for (int i = 0; i < ranks.Count; i++)
            {
                List<double> tmpRes = new List<double>();
                for (int j = 0; j < 3; j++)
                {
                    tmpRes.Add(tmpCenters[i, j]);
                }
                centers.Add(ranks[i], tmpRes);
            }
            return centers;
        }
        public static Dictionary<string, List<double>> CntrMFLingVar(Dictionary<string, List<string>> valToRanks, Dictionary<string,int> uniqVal)
        {
            Dictionary<string, List<double>> membershipFunction = new Dictionary<string, List<double>>();
            Dictionary<string, int> uniqForRanks = new Dictionary<string, int>();
            int N = 0;//мощность множества
            int m = 0;//количество рангов
            foreach(var item in valToRanks)
            {
                int rankCount = 0;
                foreach(var val in item.Value)
                {
                    rankCount += uniqVal[val];
                    N += uniqVal[val];
                }
                uniqForRanks.Add(item.Key, rankCount);
                m++;
            }
            Dictionary<string, double> freq = new Dictionary<string, double>();
            foreach (var item in uniqForRanks)
            {
                freq.Add(item.Key, (double)item.Value / N);
            }
            List<double> cntrs = new List<double>();
            cntrs.Add(0.0);
            cntrs.Add(freq.ElementAt(0).Value / 2);
            for (int i=1;i< m; i++)
            {
                double tmp = cntrs.Last() + (freq.ElementAt(i - 1).Value + freq.ElementAt(i).Value) / 2;
                cntrs.Add(tmp);
                membershipFunction.Add(freq.ElementAt(i - 1).Key, new List<double>(cntrs));
                cntrs.Remove(cntrs.First());
            }
            cntrs.Add(1.0);
            membershipFunction.Add(freq.ElementAt(m - 1).Key, cntrs);            
            return membershipFunction;
        }
        public static Bitmap CreatePicturePercent(Dictionary<string,Color> colors, Dictionary<string, double> probability)
        {
            List<double> percent = new List<double>();
            foreach (var item in probability)
            {
                percent.Add(item.Value);
            }
            int width = 25;
            int height = 5;
            List<int> pxls = new List<int>();
            pxls.Add((int)(percent[0] * width));
            for (int i = 1; i < percent.Count; i++)
            {
                pxls.Add((int)(pxls[i - 1] + (percent[i] * width)));
            }
            pxls.Add(width);            
            Bitmap img = new Bitmap(width, height);
            int countBorder = percent.Count();
            for (int i = 0, j = 0; i < countBorder; i++)
            {
                for (; j < pxls[i]; j++)
                {
                    for (int k = 0; k < height; k++)
                    {
                        img.SetPixel(j, k, colors.Values.ToList()[i]);
                    }
                }
            }            
            img.Save(Environment.CurrentDirectory +
                                        "\\Pictures\\" + String.Join(", ", percent.ToArray()) + ".jpg");
            return img;
        }
        public static double Info(List<double> probability)
        {
            double entropy = 0.00;
            foreach (var item in probability)
            {
                if (item > 0)
                {
                    entropy -= item * (Math.Log(item) / Math.Log(2));
                }                
            }
            return entropy;
        }
        public static Dictionary<string, double> ProbabilityOfClass(List<string> classes, List<int> indexes, string[] outputs)
        {
            Dictionary<string, double> probabilitty = new Dictionary<string, double>();
            foreach(var cls in classes)
            {
                probabilitty.Add(cls, 0);
            }
            foreach(var cls in indexes)
            {
                probabilitty[outputs[cls]] += 1;                
            }
            int countElements = indexes.Count;
            if (countElements > 0)
            {
                foreach (var item in probabilitty.Keys.ToList())
                {
                    probabilitty[item] = probabilitty[item] / countElements;
                }
            }            
            return probabilitty;
        }
        public static Dictionary<string, Dictionary<double, double>> DegreeOfMembDouble(Dictionary<string, List<double>> centers, List<double> uniqInputs)
        {
            Dictionary<string, Dictionary<double, double>> rankDegreeze = new Dictionary<string, Dictionary<double, double>>();
            foreach (var rank in centers)
            {
                Dictionary<double, double> degreeze = new Dictionary<double, double>();
                foreach (var uniqZn in uniqInputs)
                {
                    if (Convert.ToDouble(uniqZn) < rank.Value[0] || Convert.ToDouble(uniqZn) > rank.Value[2])
                    {
                        degreeze.Add(uniqZn, 0.00);
                    }
                    else
                    {
                        if (Convert.ToDouble(uniqZn) == rank.Value[1])//если равно центру
                        {
                            degreeze.Add(uniqZn, 1.00);
                        }
                        else
                        {
                            if (Convert.ToDouble(uniqZn) < rank.Value[1])
                            {
                                degreeze.Add(uniqZn,
                                    (Convert.ToDouble(uniqZn) - rank.Value[0]) / (rank.Value[1] - rank.Value[0]));
                            }
                            else
                            {
                                degreeze.Add(uniqZn,
                                    (rank.Value[2] - Convert.ToDouble(uniqZn)) / (rank.Value[2] - rank.Value[1]));
                            }
                        }
                    }
                }
                rankDegreeze.Add(rank.Key, degreeze);
            }
            return rankDegreeze;
        }
        //public static Dictionary<string, Dictionary<string, double>> DegreeOfMembLing(Dictionary<string, List<double>> centers, Dictionary<string, List<string>> valToRanks, Dictionary<string, int> uniqInputs)
        //{
        //    Dictionary<string, Dictionary<string, double>> rankDegreeze = new Dictionary<string, Dictionary<string, double>>();
        //    Dictionary<string, double> valDegreeze = new Dictionary<string, double>();
        //    foreach (var val in uniqInputs)
        //    {
        //        int indCntr = valToRanks.Keys.ToList().IndexOf(valToRanks.FirstOrDefault(x => x.Value.Contains(val.Key)).Key);

        //    }
        //    return rankDegreeze;
        //}
        public static Dictionary<string, double> Confirm (List<string> inputs)
        {
            Dictionary<string, double> result = new Dictionary<string, double>();
            foreach(var item in inputs)
            {
                result.Add(item, Convert.ToDouble(item));
            }
            return result;
        }
        public static List<double> ReflectionOnTheZeroOne (List<double> listString)
        {
            List<double> result = new List<double>();
            double a = listString.Min();
            double b = listString.Max();
            foreach(var item in listString)
            {
                double tmp = (item - a) / (b - a);
                result.Add(tmp);
            }
            return result;
        }
        public static Dictionary<string, Dictionary<double,double>> ReturnToAB (Dictionary<string,Dictionary<double,double>> degr, List<double> uniqVal)
        {
            Dictionary<string, Dictionary<double, double>> result = new Dictionary<string, Dictionary<double, double>>();
            foreach (var item in degr)
            {
                int i = 0;
                Dictionary<double, double> tmp = new Dictionary<double, double>();
                foreach(var value in item.Value)
                {
                    tmp.Add(uniqVal[i], value.Value);
                    i++;
                }
                result.Add(item.Key, tmp);
            }
            return result;
        }
    }
}

