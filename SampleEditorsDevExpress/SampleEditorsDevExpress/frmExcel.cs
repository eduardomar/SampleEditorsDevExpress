using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.XtraBars.Ribbon;
using System.Collections.Generic;
using DevExpress.Spreadsheet.Export;

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

        public Worksheet Worksheet(int index)
         {
            if(index >= 0 && index < spreadsheetControl1.Document.Worksheets.Count)
             {
                 return spreadsheetControl1.Document.Worksheets[index];
             }

            return null;
         }

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

        public void ExportToDataTable()
         {
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
            DataTable dataTable = worksheet.CreateDataTable(worksheet.GetUsedRange(), true);
            DataTableExporter exporter = worksheet.CreateDataTableExporter(worksheet.GetUsedRange(), dataTable, true);
            exporter.Export();
         }

        public void Create()
        {
            if(dt != null)
             {
                spreadsheetControl1.Document.Styles.Add("HeaderStyle");
                Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
                
                
                //TableStyleCollection tableStyle = spreadsheetControl1.Document.TableStyles;

                worksheet.Name = dt.TableName;
                worksheet.Import(dt, AddHeader, FirstRowIndex, FirstColIndex);

                /*Table table = worksheet.Tables.Add(worksheet.Range.FromLTRB(FirstColIndex, FirstRowIndex, dt.Columns.Count - 1, dt.Rows.Count), true);
                table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleLight9];*/
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

                    /*Table table = worksheet.Tables.Add(worksheet.Range.FromLTRB(FirstColIndex, FirstRowIndex, ds.Tables[i].Columns.Count - 1, ds.Tables[i].Rows.Count), true);
                    table.Style = spreadsheetControl1.Document.TableStyles[BuiltInTableStyleId.TableStyleLight9];*/
                 }
             }
            else if(path != null && File.Exists(path))
             {
                IWorkbook workbook = spreadsheetControl1.Document;

                using (FileStream stream = new FileStream(path, FileMode.Open))
                {
                    workbook.LoadDocument(stream, DocumentFormat.OpenXml);

                    foreach(Worksheet worksheet in workbook.Worksheets)
                     {
                        worksheet.GetUsedRange().AutoFitRows();
                        worksheet.GetUsedRange().AutoFitColumns();
                     }
                }
             }

            if (spreadsheetControl1.Document.Worksheets.Count > 0)
             {
                spreadsheetControl1.Document.Worksheets.ActiveWorksheet = spreadsheetControl1.Document.Worksheets[0];
             }
        }
    }
}
