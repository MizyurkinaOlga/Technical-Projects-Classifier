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
        public static Dictionary<string, double> CentersOfFP (string[] values)
        {
            int countVal = values.Length;
            Dictionary<string, double> membershipFunction = new Dictionary<string, double>();
            foreach(var val in values)
            {
                if (membershipFunction.ContainsKey(val))
                {
                    membershipFunction[val] += 1;
                }
                else
                {
                    membershipFunction.Add(val, 1);
                }
                
            }
            foreach(var val in membershipFunction.Keys.ToList())
            {
                membershipFunction[val] = membershipFunction[val] / countVal;
            }

            return membershipFunction;
        }
    }
}
