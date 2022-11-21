using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSS_DatalogCombiner
{
    public partial class Form1 : Form
    {
        Combiner combiner = new Combiner();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            combiner.CreateExcel();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.Filter = "csv files (*.*)|*.csv";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("File " + dialog.FileName + " is selected.");
                combiner.ImportFile(dialog.FileName);
                MessageBox.Show("File " + dialog.FileName + " is saved.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Console.WriteLine("===Path Name List===");
            combiner.DisplayPathNameHandler();
            Console.WriteLine("===End of List===");
        }
    }
}
