using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DecisionTree
{
    public partial class SortRanks : Form
    {
        public SortRanks(string attributeName, string[] attributeVal)
        {
            List<string> uniqVal = new List<string>();
            foreach(var val in attributeVal)
            {
                if (!uniqVal.Contains(val))
                {
                    uniqVal.Add(val);
                }
            }
            InitializeComponent(attributeName, uniqVal);
        }
        private void buttonCancel_click(object sender, EventArgs e)
        {

        }
        private void buttonOK_click(object sender, EventArgs e)
        {

        }
    }
}
