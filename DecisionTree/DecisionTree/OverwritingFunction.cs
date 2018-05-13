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
    public partial class OverwritingFunction : Form
    {
        public OverwritingFunction(string attributeName)
        {
            InitializeComponent();
            label2.Text = attributeName;
        }

        private void OverwritingFunction_Load(object sender, EventArgs e)
        {

        }

    }
}
