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
        //public static Dictionary<string, List<double>> CntrMFLingVar (Dictionary<string, int> uniqValues, int countValAll)
        //{
        //    Dictionary<string, List<double>> membershipFunction = new Dictionary<string, List<double>>();

        //    foreach(var val in uniqValues.Keys.ToList())
        //    {
        //        List<double> tmp = new List<double>();
        //        tmp.Add((double)uniqValues[val] / countValAll);
        //        membershipFunction.Add(val, tmp);
        //    }
        //    double cPred = 0;
        //    double fPred = 0;
        //    double fNow = 0;
        //    foreach (var val in membershipFunction.Keys.ToList())
        //    {
        //        fNow = membershipFunction[val];
        //        membershipFunction[val] = cPred + (fPred + fNow) / 2;
        //        fPred = fNow;
        //        cPred = membershipFunction[val];
        //    }
        //    return membershipFunction;
        //}
        public static Dictionary<string, List<double>> CntrMFUniCover(List<string> ranks, string[] values)
        {
            Dictionary<string, List<double>> centers = new Dictionary<string, List<double>>();

            List<double> valDouble = new List<double>();
            foreach (string item in values)
            {
                valDouble.Add(Convert.ToDouble(item));
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
        public static Bitmap CreatePicturePercent(List<Color> colors, List<double> percent)
        {
            int width = 50;
            int height = 10;
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
                        img.SetPixel(j, k, colors[i]);
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
                entropy -= item * (Math.Log(item) / Math.Log(2));
            }
            return entropy;
        }
        public static Dictionary<string, double> ProbabilityOfClass(string[] outputs)
        {
            Dictionary<string, double> probabilitty = new Dictionary<string, double>();
            foreach(var cls in outputs)
            {
                if (probabilitty.ContainsKey(cls))
                {
                    probabilitty[cls] += 1;
                }
                else
                {
                    probabilitty.Add(cls, 1);
                }
            }
            int countElements = outputs.Length;
            foreach(var item in probabilitty.Keys.ToList())
            {
                probabilitty[item] = probabilitty[item] / countElements;
            }
            return probabilitty;
        }
            
    }
}
