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
    public partial class ExpertReview : Form
    {
        Dictionary<string, List<string>> valToRanks;
        int indexRnks = 0;
        List<string> uniqVal;
        List<string> ranks;
        public ExpertReview(string attrbNm, List<string> rnks, Dictionary<string,int> values)
        {
            uniqVal = new List<string>();
            ranks = new List<string>(rnks);
            valToRanks = new Dictionary<string, List<string>>();
            foreach (var val in values)
            {
                uniqVal.Add(val.Key);
            }
            InitializeComponent();
            atributeName.Text = attrbNm;
            rankName.Text = rnks[indexRnks];
            foreach(var item in uniqVal)
            {
                checkedListBox1.Items.Add(item);
            }
        }
        private void btnCheck_Click(object sender, EventArgs e)
        {
            List<string> tmpCheck = new List<string>();
            foreach(var item in checkedListBox1.CheckedItems)
            {
                tmpCheck.Add(item.ToString());
                uniqVal.Remove(item.ToString());
            }
            valToRanks.Add(ranks[indexRnks], tmpCheck);
            checkedListBox1.Items.Clear();
            indexRnks++;
            rankName.Text = ranks[indexRnks];
            foreach (var item in uniqVal)
            {
                checkedListBox1.Items.Add(item);
            }
            if (indexRnks == ranks.Count - 1)
            {
                btnCheck.Visible = false;
                btnCheck.Enabled = false;
                button1.Visible = true;
                button1.Enabled = true;                
            }
        }
        public Dictionary<string, List<string>> RetValToRanks()
        {
            return valToRanks;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> tmpCheck = new List<string>();
            foreach (var item in checkedListBox1.CheckedItems)
            {
                tmpCheck.Add(item.ToString());
                uniqVal.Remove(item.ToString());
            }
            valToRanks.Add(ranks[indexRnks], tmpCheck);
            checkedListBox1.Items.Clear();
        }
    }
}
