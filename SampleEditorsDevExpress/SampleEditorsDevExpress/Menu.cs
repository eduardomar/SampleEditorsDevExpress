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
            if(radioGroup1.SelectedIndex == 0)
             {
                DataTable sourceTable = new DataTable("Products");
                sourceTable.Columns.Add("Product", typeof(String));
                sourceTable.Columns.Add("Price", typeof(Single));
                sourceTable.Columns.Add("Quantity", typeof(Int32));
                sourceTable.Columns.Add("Discount", typeof(Single));

                sourceTable.Rows.Add("Chocolade", 5, 15, 0.03);
                sourceTable.Rows.Add("Konbu", 9, 55, 0.1);
                sourceTable.Rows.Add("Geitost", 15, 70, 0.07);

                frmExcel frmExcel = new frmExcel(sourceTable);
                frmExcel.AddHeader = true;
                frmExcel.FirstColIndex = 0;
                frmExcel.FirstRowIndex = 0;
                frmExcel.Show();
             }
            else if(radioGroup1.SelectedIndex == 1)
             {
                DataTable sourceTable = new DataTable("Products");
                sourceTable.Columns.Add("Product", typeof(String));
                sourceTable.Columns.Add("Price", typeof(Single));
                sourceTable.Columns.Add("Quantity", typeof(Int32));
                sourceTable.Columns.Add("Discount", typeof(Single));

                sourceTable.Rows.Add("Chocolade", 5, 15, 0.03);
                sourceTable.Rows.Add("Konbu", 9, 55, 0.1);
                sourceTable.Rows.Add("Geitost", 15, 70, 0.07);

                DataSet ds = new DataSet();
                ds.Tables.Add(sourceTable);

                sourceTable = new DataTable("Names");
                sourceTable.Columns.Add("Name", typeof(String));
                sourceTable.Columns.Add("Age", typeof(Int32));

                sourceTable.Rows.Add("Eduardo", 24);
                sourceTable.Rows.Add("Omar", 24);
                sourceTable.Rows.Add("Lupe", 23);
                ds.Tables.Add(sourceTable);

                frmExcel frmExcel = new frmExcel(ds);
                frmExcel.AddHeader = true;
                frmExcel.FirstColIndex = 0;
                frmExcel.FirstRowIndex = 0;
                frmExcel.Show();
             }
            else if(radioGroup1.SelectedIndex == 2)
             {
                if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                 {
                    frmExcel frmExcel = new frmExcel(openFileDialog1.FileName);
                    frmExcel.AddHeader = true;
                    frmExcel.FirstColIndex = 0;
                    frmExcel.FirstRowIndex = 0;
                    frmExcel.Show();
                 }
             }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            Form frmWord = new frmWord();
            frmWord.Show();
        }

        private void radioGroup1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnExcel.Enabled = true;
        }
    }
}
