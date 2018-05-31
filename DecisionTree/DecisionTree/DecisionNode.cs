using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DecisionTree
{
    class DecisionNode : TreeNode
    {
        string atribute;
        string fuzzyValue;
        List<int> listIndexElement;
        double entropy;
        Dictionary<string, double> probabilityClasses;
        public DecisionNode(string atr, string rank, List<int> indexes, double info, int imgInd, Dictionary<string, double> probab)
        {
            atribute = atr;
            fuzzyValue = rank;
            listIndexElement = indexes;
            entropy = info;
            probabilityClasses = probab;
            this.Text = atr + ": " + rank;
            this.ImageIndex = imgInd;
            this.SelectedImageIndex = imgInd;
        }
        public string Atribute
        {
            get
            {
                return atribute;
            }
            set
            {
                atribute = value;
            }
        }
        public string FuzzyValues
        {
            get
            {
                return fuzzyValue;
            }
            set
            {
                fuzzyValue = value;
            }
        }
        public double Entropy
        {
            get
            {
                return entropy;
            }
            set
            {
                entropy = value;
            }
        }
        public List<int> ListIndexElements
        {
            get
            {
                return listIndexElement;
            }
            set
            {
                listIndexElement = value;
            }
        }
        public Dictionary<string, double> ProbabilityClasses
        {
            get
            {
                return probabilityClasses;
            }
            set
            {
                probabilityClasses = value;
            }
        }

    }
}
