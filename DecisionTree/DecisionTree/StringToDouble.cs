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
    public partial class StringToDouble : Form
    {
        Dictionary<string, double> conformity;
        public StringToDouble(List<string> value)
        {
            InitializeComponent();
            foreach(var item in value)
            {
                dataGridView1.Rows.Add(item, "");
            }
            conformity = new Dictionary<string, double>();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {

                conformity.Add(dataGridView1.Rows[i].Cells[0].Value.ToString(), Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value.ToString()));
            }
        }
        public Dictionary<string, double> Identify()
        {
            return conformity;
        }
    }
}
