using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraBars.Ribbon;
using System.Collections.Generic;

namespace SampleEditorsDevExpress
{
    public partial class frmExcel : RibbonForm
     {
        private String path;
        private DataTable dt;
        private DataSet ds;
        public bool AddHeader { get; set; }
        public int FirstRowIndex { get; set; }
        public int FirstColIndex { get; set; }

        private void Initialize(String path, DataTable dt, DataSet ds)
         {
            InitializeComponent();

            this.path = path;
            this.dt = dt;
            this.ds = ds;

            this.AddHeader = true;
            this.FirstColIndex = 0;
            this.FirstRowIndex = 0;
         }

        public frmExcel(String path)
         {
            Initialize(path, null, null);
         }

        public frmExcel(DataTable dt)
         {
            Initialize(null, dt, null);
         }

        public frmExcel(DataSet ds)
         {
            Initialize(null, null, ds);
         }

        private void frmExcel_Shown(object sender, EventArgs e)
        {
            if(dt != null)
             {
                Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
                //TableStyleCollection tableStyle = spreadsheetControl1.Document.TableStyles;

                worksheet.Name = dt.TableName;
                worksheet.Import(dt, AddHeader, FirstRowIndex, FirstColIndex);

                Table table = worksheet.Tables.Add(worksheet["A1:B4"], true);
                // Format the table by applying a built-in table style.
                table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleDark1];
             }
            else if(ds != null && ds.Tables.Count > 0)
             {
                WorksheetCollection worksheetCollection = spreadsheetControl1.Document.Worksheets;
                Worksheet worksheet = null;

                for(int i = 0; i < ds.Tables.Count; i++)
                 {
                    worksheet = null;

                    if (i < worksheetCollection.Count)
                     {
                        worksheet = worksheetCollection[i];
                     }
                    else
                     {
                        worksheet = worksheetCollection.Add();
                     }

                    worksheet.Name = ds.Tables[i].TableName;
                    worksheet.Import(ds.Tables[i], AddHeader, FirstRowIndex, FirstColIndex);
                 }
             }
            else if(path != null && File.Exists(path))
             {
                IWorkbook workbook = spreadsheetControl1.Document;

                using (FileStream stream = new FileStream(path, FileMode.Open))
                {
                    workbook.LoadDocument(stream, DocumentFormat.OpenXml);
                }
             }

            if (spreadsheetControl1.Document.Worksheets.Count > 0)
             {
                spreadsheetControl1.Document.Worksheets.ActiveWorksheet = spreadsheetControl1.Document.Worksheets[0];
             }
        }
    }
}
