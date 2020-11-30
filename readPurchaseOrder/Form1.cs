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
            sqlCon = new SqlConnection(conStrin);
            sqlCon.Open();
        }

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

                    Listtext.Add(item);
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
                this.lstText.Items.Clear();
                foreach (var item in pdfText(dlg.FileName))
                {

                 this.lstText.Items.Add(item);
                }
            }

            }
            catch (Exception ex)
            {

             
            }
        }
        #region exports

        private void ExportPDFToExcel()
        {
            StringBuilder text = new StringBuilder();
            if (dlg.FileName != null)
            {
                PdfReader pdfReader = new PdfReader(dlg.FileName);
            for (int page = 1; page <= pdfReader.NumberOfPages; page++)
            {
                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.UTF8.GetBytes(currentText)));
                text.Append(currentText);
                pdfReader.Close();
            }
            //Response.Clear();
            //Response.Buffer = true;
            //Response.AddHeader("content-disposition", "attachment;filename=ReceiptExport.xls");
            //Response.Charset = "";
            //Response.ContentType = "application/vnd.ms-excel";
            //Response.Write(text);
            //Response.Flush();
            //Response.End();
            }
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
    }
}
