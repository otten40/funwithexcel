using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace funwithexcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(@"C:\Users\320051\Desktop\Copy_EAM_WIP_v7.xlsx", 15);
            MessageBox.Show("Rows: " + excel.LastRow().ToString() +
                            "\nrReadCell Test: " + excel.ReadCell(5,1));


            MessageBox.Show(excel.FindValue(Input.Text));
        }
    }
}
