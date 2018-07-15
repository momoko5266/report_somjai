using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using DGVPrinterHelper;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Threading;
using System.Globalization;

namespace ReportSomjai
{
    public partial class Report : Form
    {
        public Report()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string date_start = dateTimePicker1.Value.ToString("M/d/yyyy");
            string date_end = dateTimePicker2.Value.ToString("M/d/yyyy");


        }





        private void btn_search_Click(object sender, EventArgs e)
        {


            dataGridView1.ClearSelection();
            this.load_data();

            // this.load_data();
        }
         

        private void sum_data ()
        {
            decimal Total = 0;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                Total += Convert.ToDecimal(dataGridView1.Rows[i].Cells[6].Value);
            }

           
            int rowIdx = dataGridView1.Rows.Count - 1;
            dataGridView1.Rows[rowIdx].Cells[6].Value = Total;
           
        }

        private void sum_min()
        {
            decimal cv = 0;
            decimal sum = 0;
            decimal Total = 0;
                        
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
               

                
                         if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "ใบลดหนี้")
                         {
                                cv = Convert.ToDecimal(dataGridView1.Rows[i].Cells[6].Value )* -1;
                                  
                         }
                   
                
                
                
              
            }
              int rowIdx = dataGridView1.Rows.Count - 1;
            dataGridView1.Rows[rowIdx].Cells[6].Value = cv;




        }

        private void load_data()
        {
            try
            {
               if (comboBox1.Text == "รายงานเจ้าหนี้")
                {

                    string sql = "SELECT c.chqNo,c.supplierId,iv.involvedPartyName,c.chqDate,sb.billDocId,si.deliveryDocId,bc.branchName,ap.paymentAmount,si.netAmount,ap.paymentMethodDesc FROM ChqPays c INNER JOIN ApBillpaymentMethods ap ON  ap.apBillpaymentMethodId=c.apBillpaymentMethodId  INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId";
                    DbClass objproduct = new DbClass();
                    DataTable product_dt = objproduct.GetData(sql, "barnd_dt");

                    dataGridView1.DataSource = product_dt;
                    
                    this.dataGridView1.Columns[8].DefaultCellStyle.Format = "n2";
                    this.dataGridView1.Columns[7].DefaultCellStyle.Format = "n2";
                    this.dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    this.dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    dataGridView1.Columns[0].HeaderCell.Value = "เลขที่เช็ค";
                    dataGridView1.Columns[1].HeaderCell.Value = "รหัสซัพพลายเออ";
                    dataGridView1.Columns[2].HeaderCell.Value = "ชื่อซัพพลายเออร์";
                    //dataGridView1.Columns["เลขที่เช็ค"].CellTemplate.ValueType = typeof(String);
                    dataGridView1.Columns[3].HeaderCell.Value = "วันที่เช็ค";
                    //   dataGridView1.Columns[2].CellTemplate.ValueType = typeof(string);
                    dataGridView1.Columns[4].HeaderCell.Value = "เลขที่วางบิล";
                    // dataGridView1.Columns[3].CellTemplate.ValueType = typeof(string);
                    dataGridView1.Columns[5].HeaderCell.Value = "เลขที่ใบส่งสินค้า";
                    dataGridView1.Columns[6].HeaderCell.Value = "ชื่อสาขา";
                    dataGridView1.Columns[7].HeaderCell.Value = "จำนวนเงินจ่าย";
                    dataGridView1.Columns[8].HeaderCell.Value = "จำนวนเงิน";
                    dataGridView1.Columns[9].HeaderCell.Value = "ประเภทการจ่ายเงิน";
                    //sum_min();

                }
                else if (comboBox1.Text == "รายงานการจ่ายเงิน")
                {
                    string sql = "SELECT chqNo,c.chqDate,si.netAmount,iv.involvedPartyName,bc.branchName,c.chqPrintBy FROM ChqPays c INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.chqPayId INNER JOIN ApBillpaymentMethods ap ON ap.apBillpaymentMethodId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId ";
                    DbClass objproduct = new DbClass();
                    DataTable product_dt = objproduct.GetData(sql, "barnd_dt");


                    dataGridView1.DataSource = product_dt;

                    this.dataGridView1.Columns[2].DefaultCellStyle.Format = "n2";
                    this.dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns[0].HeaderCell.Value = "เลขที่เช็ค";
                    // dataGridView1.Columns[0].ValueType = typeof(String);
                    dataGridView1.Columns[1].HeaderCell.Value = "วันที่เช็ค";
                    //  dataGridView1.Columns[1].ValueType = typeof(String);
                    dataGridView1.Columns[2].HeaderCell.Value = "จำนวนเงิน";
                    dataGridView1.Columns[3].HeaderCell.Value = "นามเช็ค";
                    dataGridView1.Columns[4].HeaderCell.Value = "ชื่อซัพพลายเออร์";
                    dataGridView1.Columns[5].HeaderCell.Value = "ผู้บันทึกรายการ";
                }else if (comboBox1.Text =="รายงานลูกหนี้" )
                {
                    string sql = "SELECT  st.documentId, cb.customerBillId,br.branchName,cb.documentId,ct.documentId,ct.customerName,spt.paymentAmount,ti.documentId,cr.receiptNumber,cpm.paymentDetails FROM CustomerPaymentReceipts cr INNER JOIN CustomerBills cb ON cb.customerId = cr.customerId INNER JOIN CustomerCreditNotes ct ON ct.customerBillId = cb.customerBillId INNER JOIN CustomerPaymentMethods cpm ON cpm.customerPaymentId = cr.customerPaymentId INNER JOIN CustomerPayments cp ON cp.customerPaymentId =  cr.customerPaymentId INNER JOIN SaleTransactions st ON st.customerBillId = cb.customerBillId INNER JOIN SaleTransactionPayments spt ON spt.customerPaymentReceiptId=cr.customerPaymentReceiptId INNER JOIN TaxInvoices ti ON ti.taxInvoiceId = cr.taxInvoiceId INNER JOIN BranchConfig br ON br.branchId = cb.branchId WHERE cpm.paymentDetails IS NOT NULL ";
                    DbClass objproduct = new DbClass();
                    DataTable product_dt = objproduct.GetData(sql, "barnd_dt");


                    dataGridView1.DataSource = product_dt;
                    this.dataGridView1.Columns[6].DefaultCellStyle.Format = "n2";
                    dataGridView1.Columns[0].HeaderCell.Value = "เลขที่ใบอินวอช";
                    // dataGridView1.Columns[0].ValueType = typeof(String);
                    dataGridView1.Columns[1].HeaderCell.Value = "รหัสลูกค้า";
                    //  dataGridView1.Columns[1].ValueType = typeof(String);
                    dataGridView1.Columns[2].HeaderCell.Value = "ชื่อสาขา";
                    dataGridView1.Columns[3].HeaderCell.Value = "เลขที่บิล";
                    dataGridView1.Columns[4].HeaderCell.Value = "เลขที่เครดิตโน๊ต";
                    dataGridView1.Columns[5].HeaderCell.Value = "ชื่อลูกค้า";
                    dataGridView1.Columns[6].HeaderCell.Value = "ราคา";
                    dataGridView1.Columns[7].HeaderCell.Value = "เลขที่ใบกำกับ";
                    dataGridView1.Columns[8].HeaderCell.Value = "เลขที่ใบ";
                    dataGridView1.Columns[9].HeaderCell.Value = "เลขที่ใบลดหนี้/เครดิตโน๊ต";
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


      
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            
        }



        private void button1_Click(object sender, EventArgs e)
        {
            
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try 
            {
                // instantiating the excel application class
                object misValue1 = System.Reflection.Missing.Value;
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook currentWorkbook = excelApp.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet currentWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)currentWorkbook.ActiveSheet;
                currentWorksheet.Columns.ColumnWidth = 18;
                if (dataGridView1.Rows.Count > 0)
                {
                    currentWorksheet.Cells[1, 1] = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString();
                    int i = 1;
                    foreach (DataGridViewColumn dgviewColumn in dataGridView1.Columns)
                    {
                        // Excel work sheet indexing starts with 1
                        currentWorksheet.Cells[2, i] = dgviewColumn.Name;
                        ++i;
                    }
                    Microsoft.Office.Interop.Excel.Range headerColumnRange = currentWorksheet.get_Range("A2", "AY2");
                    headerColumnRange.Font.Bold = true;
                    headerColumnRange.Font.Color = 0xFF0000;
                    int rowIndex = 0;
                    
                    for (rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
                    {
                        DataGridViewRow dgRow =  dataGridView1.Rows[rowIndex];
                        for (int cellIndex = 0; cellIndex < dgRow.Cells.Count; cellIndex++)
                        {
                            currentWorksheet.Cells[rowIndex + 3, cellIndex + 1] ="'" + dgRow.Cells[cellIndex].Value ;
                            
                        }
                    }
                    Microsoft.Office.Interop.Excel.Range fullTextRange = currentWorksheet.get_Range( "A2", "AY" + (rowIndex + 1).ToString());
                    fullTextRange.WrapText = true;
                    fullTextRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }

                using (SaveFileDialog exportSaveFileDialog = new SaveFileDialog())
                {
                    exportSaveFileDialog.Title = "Save as";
                    exportSaveFileDialog.Filter = "Microsoft Office Excel Workbook(*.xls)|*.xls";

                    if (DialogResult.OK == exportSaveFileDialog.ShowDialog())
                    {
                        string fullFileName = exportSaveFileDialog.FileName;

                        currentWorkbook.SaveAs(fullFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, System.Reflection.Missing.Value, misValue1, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, misValue1, misValue1, misValue1);
                        currentWorkbook.Saved = true;
                        MessageBox.Show("The export was successful", "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
            }

            


            //copyAlltoClipboard();
            //UpdateFont();
            //Excel.Application xlexcel;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;




            //object misValue = System.Reflection.Missing.Value;

            //xlexcel = new Excel.Application();
            //xlexcel.Visible = true;
            //xlWorkBook = xlexcel.Workbooks.Add(misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];

            //CR.Select();
            //xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }
    



                private void UpdateFont()
                {
                        //Change cell font
                    foreach (DataGridViewColumn c in dataGridView1.Columns)
                    {
                         c.DefaultCellStyle.Font = new Font("Angsana New (หัวเรื่องแบบ CS)", 15F, GraphicsUnit.Pixel);
                    }
                }

      



        private void btnok_Click(object sender, EventArgs e)
        {
            try { 
            if (comboBox1.Text == "รายงานการจ่ายเงิน")
            { 
                if (textBox1.Text == "")
                {
                    string dt1 = dateTimePicker1.Value.ToString("M/d/yyyy");
                    string dt2 = dateTimePicker2.Value.ToString("M/d/yyyy");
                    string sql = "SELECT ap.paymentMethodDesc,chqNo,c.chqDate,si.netAmount,iv.involvedPartyName,bc.branchName,c.chqPrintBy FROM ChqPays c INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.chqPayId INNER JOIN ApBillpaymentMethods ap ON ap.apBillpaymentMethodId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId WHERE c.chqDate BETWEEN '" + dt1+ "' AND '" + dt2 + "' ";
                    DbClass objproduct = new DbClass();
                    DataTable product_dt = objproduct.GetData(sql, "barnd_dt");
                    
                    dataGridView1.DataSource = product_dt;
                    this.dataGridView1.Rows.Insert(0, product_dt);
                    this.dataGridView1.Columns[2].DefaultCellStyle.Format = "n2";
                    this.dataGridView1.Columns[0].DefaultCellStyle.Format = "/t";
                    this.dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    
                    dataGridView1.Columns[0].HeaderCell.Value = "เลขที่เช็ค";
                   // dataGridView1.Columns[0].ValueType = typeof(String);
                    dataGridView1.Columns[1].HeaderCell.Value = "วันที่เช็ค";
                   // dataGridView1.Columns[1].ValueType = typeof(String);
                    dataGridView1.Columns[2].HeaderCell.Value = "จำนวนเงิน";  
                    dataGridView1.Columns[4].HeaderCell.Value = "ชื่อซัพพลายเออร์";
                    dataGridView1.Columns[5].HeaderCell.Value = "ผู้บันทึกรายการ";
                        
                }
                else
                {
                    try { 
                            string sql = "SELECT chqNo,c.chqDate,si.netAmount,iv.involvedPartyName,bc.branchName,c.chqPrintBy FROM ChqPays c INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.chqPayId INNER JOIN ApBillpaymentMethods ap ON ap.apBillpaymentMethodId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId WHERE c.chqNo LIKE '" + textBox1.Text +"%' OR involvedPartyName LIKE '" + textBox1.Text + "%'  ";
                            DbClass objproduct = new DbClass();
                            DataTable product_dt = objproduct.GetData(sql, "barnd_dt");
                    
                            dataGridView1.DataSource = product_dt;
                            this.dataGridView1.Columns[2].DefaultCellStyle.Format = "n2";
                            this.dataGridView1.Columns[1].DefaultCellStyle.Format = "n2";
                            this.dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                             //  this.dataGridView1.Columns[0].ValueType = typeof(String);

                            dataGridView1.Columns[0].HeaderCell.Value = "เลขที่เช็ค";
                         // dataGridView1.Columns[0].ValueType = typeof(String);
                            dataGridView1.Columns[1].HeaderCell.Value = "วันที่เช็ค";
                         //  dataGridView1.Columns[1].ValueType = typeof(String);
                            dataGridView1.Columns[2].HeaderCell.Value = "จำนวนเงิน";
                            dataGridView1.Columns[3].HeaderCell.Value = "นามเช็ค";
                            dataGridView1.Columns[4].HeaderCell.Value = "ชื่อซัพพลายเออร์";
                            dataGridView1.Columns[5].HeaderCell.Value = "ผู้บันทึกรายการ";
                            
                    }
                    catch
                    {
                        MessageBox.Show("กรอกข้อความไม่ถูกต้อง");
                    }
                }
            }
            else if (comboBox1.Text == "รายงานเจ้าหนี้")
                 

            {
                if (textBox1.Text == "")
                {
                    string dt1 = dateTimePicker1.Value.ToString("M/d/yyyy");
                    string dt2 = dateTimePicker2.Value.ToString("M/d/yyyy");
                    string sql = "SELECT ap.paymentMethodDesc,c.chqNo,c.chqDate,sb.billDocId,si.deliveryDocId,bc.branchName,ap.paymentAmount FROM ChqPays c INNER JOIN ApBillpaymentMethods ap ON  ap.apBillpaymentMethodId=c.apBillpaymentMethodId  INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId WHERE  c.chqDate BETWEEN '" + dt1 + "' AND '" + dt2 + "' ";
                    DbClass objproduct = new DbClass();
                    DataTable product_dt = objproduct.GetData(sql, "barnd_dt");


                    dataGridView1.DataSource = product_dt;
                        
                        this.dataGridView1.Columns[6].DefaultCellStyle.Format = "n2";
                    this.dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                  //  this.dataGridView1.Columns[1].ValueType = typeof(String);
                   // this.dataGridView1.Columns[3].ValueType = typeof(String);
                   // this.dataGridView1.Columns[4].ValueType = typeof(String);
                    dataGridView1.Columns[0].HeaderCell.Value = "ประเภทการจ่าย";
                    dataGridView1.Columns[1].HeaderCell.Value = "เลขที่เช็ค";
                   // dataGridView1.Columns[1].ValueType = typeof(String);
                    dataGridView1.Columns[2].HeaderCell.Value = "วันที่เช็ค";
                 //   dataGridView1.Columns[2].ValueType = typeof(String);
                    dataGridView1.Columns[3].HeaderCell.Value = "เลขที่วางบิล";
                 //   dataGridView1.Columns[3].ValueType = typeof(String);
                    dataGridView1.Columns[4].HeaderCell.Value = "เลขที่ใบส่งสินค้า";
                 //   dataGridView1.Columns[4].ValueType = typeof(String);
                    dataGridView1.Columns[5].HeaderCell.Value = "ชื่อสาขา";
                    dataGridView1.Columns[6].HeaderCell.Value = "จำนวนเงิน";
                        
                    }
                else
                {
                        try
                        {
                            string sql = "SELECT ap.paymentMethodDesc,c.chqNo,c.chqDate,sb.billDocId,si.deliveryDocId,bc.branchName,ap.paymentAmount FROM ChqPays c INNER JOIN ApBillpaymentMethods ap ON  ap.apBillpaymentMethodId=c.apBillpaymentMethodId  INNER JOIN InvolvedParties iv ON iv.involvedPartyId = c.supplierId INNER JOIN SupplierAppBills sb ON sb.supplierAppBillId = c.apBillpaymentMethodId INNER JOIN SupplierInvoices si ON si.supplierAppBillId = sb.supplierAppBillId INNER JOIN BranchConfig bc ON bc.branchId = sb.branchId WHERE c.chqNo LIKE '" + textBox1.Text + "%'  ";
                            DbClass objproduct = new DbClass();
                            DataTable product_dt = objproduct.GetData(sql, "barnd_dt");


                            dataGridView1.DataSource = product_dt;

                            this.dataGridView1.Columns[6].DefaultCellStyle.Format = "n2";
                            this.dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                            //  this.dataGridView1.Columns[1].ValueType = typeof(String);
                            //  this.dataGridView1.Columns[3].ValueType = typeof(String);
                            //   this.dataGridView1.Columns[4].ValueType = typeof(String);
                            dataGridView1.Columns[0].HeaderCell.Value = "ประเภทการจ่าย";
                            dataGridView1.Columns[1].HeaderCell.Value = "เลขที่เช็ค";
                            //   dataGridView1.Columns[1].ValueType = typeof(String);
                            dataGridView1.Columns[2].HeaderCell.Value = "วันที่เช็ค";
                            // dataGridView1.Columns[2].ValueType = typeof(String);
                            dataGridView1.Columns[3].HeaderCell.Value = "เลขที่วางบิล";
                            // dataGridView1.Columns[3].ValueType = typeof(String);
                            dataGridView1.Columns[4].HeaderCell.Value = "เลขที่ใบส่งสินค้า";
                            //  dataGridView1.Columns[4].ValueType = typeof(String);
                            dataGridView1.Columns[5].HeaderCell.Value = "ชื่อสาขา";
                            dataGridView1.Columns[6].HeaderCell.Value = "จำนวนเงิน";




                        }

                        catch
                        {
                            MessageBox.Show("กรอกข้อความไม่ถูกต้อง");
                        }

                }




            }
            }catch
            {
                MessageBox.Show("Error Load Data!");
            }
        }

        private void dataGridView1_CellValidating(object sender,
    DataGridViewCellValidatingEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].ErrorText = "";
            int newInteger;

            // Don't try to validate the 'new row' until finished 
            // editing since there
            // is not any point in validating its initial value.
            if (dataGridView1.Rows[e.RowIndex].IsNewRow) { return; }
            if (!int.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0)
            {
                e.Cancel = true;
                dataGridView1.Rows[e.RowIndex].ErrorText = "the value must be a non-negative integer";
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (comboBox1.Text == "รายงานเจ้าหนี้")
            { 
                if (e.RowIndex >= 0)
                 {
                     //gets a collection that contains all the rows
                     DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                     //populate the textbox from specific value of the coordinates of column and row.
                     textBox2.Text = row.Cells[1].Value.ToString();
                      textBox3.Text = row.Cells[2].Value.ToString();
               
                }
            }
            else if (comboBox1.Text == "รายงานการจ่ายเงิน")
            {
                if (e.RowIndex >= 0)
                {
                    //gets a collection that contains all the rows
                    DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                    //populate the textbox from specific value of the coordinates of column and row.
                    textBox2.Text = row.Cells[0].Value.ToString();
                    textBox3.Text = row.Cells[4].Value.ToString();

                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.linkLabel1.LinkVisited = true;

            // Navigate to a URL.
            System.Diagnostics.Process.Start("http://192.168.0.3/conx/#/home");
        }

       
    }
}
