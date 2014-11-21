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
    public partial class frmWord : Form
    {
        public class Employee
        {
            public Employee()
            { }
            private string _Nombre;
            public string Nombre
            {
                get { return _Nombre; }
                set { _Nombre = value; }
            }
            private string _Direccion;
            public string Direccion
            {
                get { return _Direccion; }
                set { _Direccion = value; }
            }

        }
        public frmWord()
        {
            InitializeComponent();

            List<Employee> dataSourece = new List<Employee>();
            dataSourece.Add(new Employee() { Nombre = "John", Direccion = "1st street"});
            dataSourece.Add(new Employee() { Nombre = "Steve", Direccion = "7th Ave." });
            
            richEditControl1.Options.MailMerge.DataSource = dataSourece;
            dataNavigator1.DataSource = dataSourece; 
        }

        private void frmWord_Load(object sender, EventArgs e)
        {

        }
    }
}
