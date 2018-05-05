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
        List<string> orderedValues;
        Dictionary<string, System.Windows.Forms.Label> labels = new Dictionary<string, System.Windows.Forms.Label>();
        Dictionary<string, System.Windows.Forms.ComboBox> order = new Dictionary<string, System.Windows.Forms.ComboBox>();

        public SortRanks(string attributeName, string[] uniqVal)
        {
            orderedValues = new List<string>();
            InitializeComponent(attributeName, uniqVal.ToList());
        }
        private void buttonCancel_click(object sender, EventArgs e)
        {

        }
        private void buttonOK_click(object sender, EventArgs e)
        {

        }
        public List<string> OrderedValues()
        {
            List<string> retUniqOrder = new List<string>();
            int k = 0;
            while (retUniqOrder.Count() < labels.Count())
            {
                int count = order.Count();
                for (int i = 0; i < count; i++)
                {
                    
                    if (order[("comboBox" + (i + 3)).ToString()].SelectedIndex == k)
                    {
                        retUniqOrder.Add(labels[("label" + (i + 3)).ToString()].Text);
                    }
                }
                k++;
            }
            return retUniqOrder;
        }

    }
}
