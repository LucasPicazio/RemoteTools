using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp6
{
    public partial class Form2 : Form
    {
        public string BRX = "";
        public Form2()
        {
            InitializeComponent();
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            BRX = checkBox1.Text;
            this.Close();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            BRX = checkBox2.Text;
            this.Close();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            BRX = checkBox3.Text;
            this.Close();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            BRX = checkBox4.Text;
            this.Close();
        }
    }
}
