using ClosedXML.Excel;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace readPurchaseOrder
{
    public partial class Form1 : Form
    {
        OpenFileDialog dlg = new OpenFileDialog();
        string conStrin = "Data Source=DELL\\SQLSERVER;Initial Catalog=Employees;User ID=sa;Password=Sql2008";
        SqlConnection sqlCon;
        SqlCommand sqlCmd;
        string EmployeeId = "";


        public Form1()
        {
            InitializeComponent();
            //sqlCon = new SqlConnection(conStrin);
            //sqlCon.Open();
        }
        DataTable Encabezado = new DataTable();
        DataTable Contenido = new DataTable();
        DataTable TotalesLineTotal = new DataTable();
        DataTable TotalesOrderTotal = new DataTable();


        #region pdf 

        public static List<string> pdfText(string path)
        {
            List<string> Listtext = new List<string>();
            try
            {

            PdfReader reader = new PdfReader(path);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text = PdfTextExtractor.GetTextFromPage(reader, page);
                
                foreach (var item in parseString(text))
                {
                        if (item.Contains("\n"))
                        {
                            char[] separator = { '\n' };
                            
                            foreach (var word in item.Split(separator))
                            {

                            Listtext.Add(word);
                            }
                        }else { Listtext.Add(item); }
                    
                }
            }
            reader.Close();
            }
            catch (Exception ex)
            {

                
            }
            return Listtext;
        }
        private static string[] parseString(string text)
        {
            char[] separator = { ' ' };
            string[] items;
            try
            {
            items=text.Split(separator);

            }
            catch (Exception ex)
            {

                items = new string[] { "" };
            }
            return items;
        }

        #region dataTables
        private void createTables()
        {
            Encabezado.Columns.Add("orderNo");
            Encabezado.Columns.Add("revisionNo");
            Encabezado.Columns.Add("vendor");
            Encabezado.Columns.Add("shipTo");
            Encabezado.Columns.Add("factory");
            Encabezado.Columns.Add("markFor");
            Encabezado.Columns.Add("billTo");
            Encabezado.Columns.Add("fiscalRep");
            Encabezado.Columns.Add("agent");
            Encabezado.Columns.Add("purchaseGroup");
            Encabezado.Columns.Add("customerPo");
            Encabezado.Columns.Add("poPrint");
            Encabezado.Columns.Add("orderType");
            Encabezado.Columns.Add("customerDept");
            Encabezado.Columns.Add("poIssue");
            Encabezado.Columns.Add("poGroup");
            Encabezado.Columns.Add("plant");
            Encabezado.Columns.Add("poContact");
            Encabezado.Columns.Add("paymentCategory");
            Encabezado.Columns.Add("mfgOrigin");
            Encabezado.Columns.Add("dateSent");
            Encabezado.Columns.Add("businnesType");
            Encabezado.Columns.Add("materialNumber");
            Encabezado.Columns.Add("poItem");
            Encabezado.Columns.Add("season");
            Encabezado.Columns.Add("incoterms");
            Encabezado.Columns.Add("contractualDeliveryDate");
            Encabezado.Columns.Add("inboundPkg");
            Encabezado.Columns.Add("incotermsPlace");
            Encabezado.Columns.Add("handoverDate");
            Encabezado.Columns.Add("mfgProcess");
            Encabezado.Columns.Add("harborPort");
            Encabezado.Columns.Add("customerHandoverPlace");
            Encabezado.Columns.Add("quality");
            Encabezado.Columns.Add("shipMode");
            Encabezado.Columns.Add("shade");
            Encabezado.Columns.Add("centralPoNumber");
            Encabezado.Columns.Add("model");
            Encabezado.Columns.Add("productType");
            Encabezado.Columns.Add("merchDivision");
            Encabezado.Columns.Add("colorDescription");
            Encabezado.Columns.Add("class");
            Encabezado.Columns.Add("conceptShortDesc");
            Encabezado.Columns.Add("fabricContent");
            Encabezado.Columns.Add("board");
            Encabezado.Columns.Add("fishWildlifeInd");
            Encabezado.Columns.Add("gender");
            Encabezado.Columns.Add("downFeatherInd");
            Encabezado.Columns.Add("pattern");
            Encabezado.Columns.Add("fixture");
            Encabezado.Columns.Add("rigIndicator");
            Encabezado.Columns.Add("fabrication");
            
            Contenido.Columns.Add("size");
            Contenido.Columns.Add("upc");
            Contenido.Columns.Add("msrp");
            Contenido.Columns.Add("customerSellingPrice");
            Contenido.Columns.Add("price");
            Contenido.Columns.Add("quantity");
            Contenido.Columns.Add("amount");

            TotalesLineTotal.Columns.Add("polLineTotalQuantity");
            TotalesLineTotal.Columns.Add("polLineTotalAmount");
            

            TotalesOrderTotal.Columns.Add("purcahseOrderQuantity");
            TotalesOrderTotal.Columns.Add("purcahseOrderAmount");
            
        }
        #endregion
        private static string GetTextFromPDF(string path)
        {
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(path))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return text.ToString();
        }
        #endregion

        private void btnLoad_Click(object sender, EventArgs e)
        {
            
            // set file filter of dialog   
            dlg.Filter = "pdf files (*.pdf) |*.pdf;";
            try
            {
            dlg.ShowDialog();
            if (dlg.FileName != null)
            {
                    lstText.DataSource = null;
                    this.lstText.Items.Clear();
                    List<string> dummy = pdfText(dlg.FileName);
                foreach (var item in pdfText(dlg.FileName))
                {

                        
                    
                 this.lstText.Items.Add(item);
                }
                    Encabezado.Rows.Add(dummy[28],dummy[31], 
                        dummy[34]+ " "+ dummy[39]+" " + dummy[40]+" "+dummy[46]
                        + " "+dummy[47]+" "+ dummy[48] + " "+ dummy[49] + " "+
                        dummy[56] + " "+ dummy[57] + " "+ dummy[58] + " "+
                        dummy[59] + " "+ dummy[60] + " "+ dummy[61] + " "+
                        dummy[65] + " "+ dummy[66] + " "+ dummy[67] + " "
                        ,
                        dummy[38]+" "+ dummy[41] + " "+ dummy[42] + " "+
                        dummy[43] + " "+ dummy[44] + " "+ dummy[45] + " "+
                        dummy[50] + " "+ dummy[51] + " "+ dummy[52] + " "+
                        dummy[53] + " "+ dummy[54] + " "+ dummy[55] + " "+
                        dummy[53] + " "+ dummy[54] + " "+ dummy[55] + " "+
                        dummy[62] + " "+ dummy[63] + " "+ dummy[64] + " "+
                        dummy[68] + " " + dummy[69],
                        dummy[71] + " " + dummy[74]+ " " + dummy[75]+
                        " " + dummy[76] + " " + dummy[77] + " " + dummy[78]+
                        " " + dummy[79]+ " " + dummy[80] + " " + dummy[81]+
                        " " + dummy[82] + " " + dummy[83] + " " + dummy[84]
                        + " " + dummy[85] + " " + dummy[86] + " " + dummy[87]
                        + " " + dummy[88] + " " + dummy[89] + " " + dummy[90]
                        + " " + dummy[91]
                        ,"",

                        dummy[94] + " " + dummy[97]+ " " + dummy[98]+
                        " " + dummy[99]+ " " + dummy[100]+
                        " " + dummy[101] + " " + dummy[102]+" "+ dummy[103]+
                        " " + dummy[104] + " " + dummy[105] + " " + dummy[106]+
                        " " + dummy[107] + " " + dummy[108] + " " + dummy[109]+
                        " "+dummy[110]+" "+ dummy[111]+" "+ dummy[113],""
                        , dummy[115] + " "+dummy[116] + " "+dummy[117] + " "+
                        dummy[118] + " "+ dummy[119] + " "+ dummy[120] + dummy[121] 

                        );

                }

            }
            catch (Exception ex)
            {

             
            }
        }
        #region exports

        private void ExportPDFToExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                workbook.SaveAs("C://TPMX//HelloWorld.xlsx");
            }
            //StringBuilder text = new StringBuilder();
            //if (dlg.FileName != null)
            //{
            //    PdfReader pdfReader = new PdfReader(dlg.FileName);
            //for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            //{
            //    ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
            //    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
            //    currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.UTF8.GetBytes(currentText)));
            //    text.Append(currentText);
            //    pdfReader.Close();
            //}
            //Response.Clear();
            //Response.Buffer = true;
            //Response.AddHeader("content-disposition", "attachment;filename=ReceiptExport.xls");
            //Response.Charset = "";
            //Response.ContentType = "application/vnd.ms-excel";
            //Response.Write(text);
            //Response.Flush();
            //Response.End();
        //}
        }

        
        #endregion
        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            ExportPDFToExcel();
        }

        #region base datos
        private void SaveData()
        {
            try
            {
                if (sqlCon.State == ConnectionState.Closed)
                {
                    sqlCon.Open();
                }
                DataTable dtData = new DataTable();
                sqlCmd = new SqlCommand("spEmployee", sqlCon);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("@ActionType", "SaveData");
                sqlCmd.Parameters.AddWithValue("@EmployeeId", EmployeeId);
                //sqlCmd.Parameters.AddWithValue("@Name", textBox1.Text);
                //sqlCmd.Parameters.AddWithValue("@City", textBox2.Text);
                //sqlCmd.Parameters.AddWithValue("@Department", textBox4.Text);
                //sqlCmd.Parameters.AddWithValue("@Gender", comboBox1.Text);
                int numRes = sqlCmd.ExecuteNonQuery();
                if (numRes > 0)
                {
                    MessageBox.Show("Data saved successfully !!!");

                }
                else MessageBox.Show("Please try again");

            }
            catch (Exception ex)
            {

                MessageBox.Show("Error:-" + ex.Message);
            }
        }
        private DataTable FetchEmpDetails()
        {
            if (sqlCon.State == ConnectionState.Closed)
            {
                sqlCon.Open();
            }
            DataTable dtData = new DataTable();
            sqlCmd = new SqlCommand("spEmployee", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.Parameters.AddWithValue("@ActionType", "FetchData");
            SqlDataAdapter sqlSda = new SqlDataAdapter(sqlCmd);
            sqlSda.Fill(dtData);
            return dtData;
        }
        private DataTable FetchEmpRecords(string empId)
        {
            if (sqlCon.State == ConnectionState.Closed)
            {
                sqlCon.Open();
            }
            DataTable dtData = new DataTable();
            sqlCmd = new SqlCommand("spEmployee", sqlCon);
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.Parameters.AddWithValue("@ActionType", "FetchRecord");
            sqlCmd.Parameters.AddWithValue("EmployeeId", empId);
            SqlDataAdapter sqlSda = new SqlDataAdapter(sqlCmd);
            sqlSda.Fill(dtData);
            return dtData;
        }
        private void Delete()
        {
            if (!string.IsNullOrEmpty(EmployeeId))
            {
                try
                {
                    if (sqlCon.State == ConnectionState.Closed)
                    {
                        sqlCon.Open();
                    }
                    DataTable dtData = new DataTable();
                    sqlCmd = new SqlCommand("spEmployee", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ActionType", "DeleteData");
                    sqlCmd.Parameters.AddWithValue("@EmployeeId", EmployeeId);
                    int numRes = sqlCmd.ExecuteNonQuery();
                    if (numRes > 0)
                    {
                        MessageBox.Show("Record deleted successfully!!!");
                        

                    }
                    else
                    {
                        MessageBox.Show("Please try Again!!!");
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("Error:-" + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please Select a record");
            }
        }

        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            createTables();
        }
    }
}
