using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                double a = Convert.ToDouble(If.Text);
                double b = Convert.ToDouble(CTr.Text);
                double c = Convert.ToDouble(Ohmm.Text);
                double d = Convert.ToDouble(Lengthm.Text);
                label1.Text = (a / b * 2).ToString()+" Rct + " + (a / b * 2 * c * d).ToString(); ;
            }
            catch (Exception)
            {
                MessageBox.Show("invalid input");
            }            
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
