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
    public partial class SelectTableSheet : Form
    {
        public SelectTableSheet(string[] tables)
        {
            InitializeComponent();
            this.listBox1.DataSource = tables;
        }
        public int Selection
        {
            get
            {
                return this.listBox1.SelectedIndex;
            //    return this.listBox1.SelectedItem as string;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
