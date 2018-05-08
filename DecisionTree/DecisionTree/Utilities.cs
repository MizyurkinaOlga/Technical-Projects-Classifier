using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MExcel = Microsoft.Office.Interop.Excel;

namespace DecisionTree
{
    public static class Utilities
    {
        public static Dictionary<string,string> TypeOfInputs(string[,] data, string[] inputs)
        {
            Dictionary<string, string> typeOfInputs = new Dictionary<string, string>();
            for (int j = 0; j < inputs.Length; j++)
            {
                double tmp;
                bool douBle = true;
                for (int i = 0; i < data.Length/inputs.Length; i++)
                {
                    if (!Double.TryParse(data[i, j], out tmp))
                    {
                        douBle = false;
                        break;
                    }
                }
                if (douBle == false)
                {
                    typeOfInputs.Add(inputs[j], "string");
                }
                else
                {
                    typeOfInputs.Add(inputs[j], "double");
                }
            }
            return typeOfInputs;
        }
        public static Dictionary<int,int> RangeOfData(MExcel.Worksheet ExcSheet)
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
            Dictionary<string, int> uniqVal = new Dictionary<string, int>();
            foreach (var val in values)
            {
                if (uniqVal.ContainsKey(val))
                {
                    uniqVal[val] += 1;
                }
                else
                {
                    uniqVal.Add(val, 1);
                }
            }
            return uniqVal;
        }        
        public static Dictionary<string, double> CentersOfFP (Dictionary<string, int> uniqValues, int countValAll)
        {
            Dictionary<string, double> membershipFunction = new Dictionary<string, double>();
            foreach(var val in uniqValues.Keys.ToList())
            {
                membershipFunction.Add(val, (double) uniqValues[val] / countValAll);
            }
            double cPred = 0;
            double fPred = 0;
            double fNow = 0;
            foreach (var val in membershipFunction.Keys.ToList())
            {
                fNow = membershipFunction[val];
                membershipFunction[val] = cPred + (fPred + fNow) / 2;
                fPred = fNow;
                cPred = membershipFunction[val];
            }
            return membershipFunction;
        }
    }
}
