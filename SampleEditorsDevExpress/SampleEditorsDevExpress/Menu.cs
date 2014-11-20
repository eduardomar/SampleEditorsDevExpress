using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SampleEditorsDevExpress
{
    public partial class Menu : Form
    {
        public Menu()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            Form frmExcel = new frmExcel();
            frmExcel.Show();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            Form frmWord = new frmWord();
            frmWord.Show();
        }
    }
}
