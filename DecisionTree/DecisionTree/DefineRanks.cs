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
    public partial class DefineRanks : Form
    {
        public DefineRanks(string attributeName)
        {
            InitializeComponent();
            label2.Text = attributeName;
        }

        public List<string> Identify()
        {
            string[] separator = new string[] { "\r\n" };
            return textBox1.Text.Split(separator, StringSplitOptions.None).ToList();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text== "Введите ранги атрибута...")
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Введите ранги атрибута...";
                textBox1.ForeColor = Color.Gray;
            }
        }
    }
}
