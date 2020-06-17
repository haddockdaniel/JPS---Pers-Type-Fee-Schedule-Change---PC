using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;
using JurisSVR;
using System.Reflection;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public string fileName = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            if (string.IsNullOrEmpty(fileName))
                MessageBox.Show("Please select an Excel file before proceeding", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                CliObj cli = null;
                List<CliObj> clients = new List<CliObj>();

                int lastUsedRow = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                int count = 1;

                string clientCode = "";
                string clientRateSched = "";

                if (xlRange.Cells[1,2] != null && xlRange.Cells[1,2].Value2 != null)
                {
                    clientCode = xlRange.Cells[1,2].Value2.ToString();
                    String ssql = "select CliFeeSch from client where dbo.jfn_FormatClientCode(clicode) = '" + clientCode + "'";
                    DataSet dd = _jurisUtility.RecordsetFromSQL(ssql);
                    clientRateSched = dd.Tables[0].Rows[0][0].ToString();
                }
                Error err = new Error();
                if (err.doesCliExist(clientCode, _jurisUtility))
                {
                    for (int i = 2; i <= lastUsedRow; i++)
                    {
                        cli = new CliObj();
                        cli.error = false;
                        cli.errorMess = "";
                        for (int j = 1; j <= 2; j++)
                        {

                            if (j == 1)
                            {
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                {
                                    cli.PT = xlRange.Cells[i, j].Value2.ToString();
                                }

                            }
                            else if (j == 2)
                            {
                                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                {
                                    if (err.isRateNumeric(xlRange.Cells[i, j].Value2.ToString(), _jurisUtility))
                                        cli.rate = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                                    else
                                    {
                                        cli.error = true;
                                        cli.errorMess = "The Rate specified: " + xlRange.Cells[i, j].Value2.ToString() + " is not numeric. Please update the spreadsheet";
                                    }
                                }

                            }

                        }

                        clients.Add(cli);
                        UpdateStatus("Accessing Spreadsheet", count, lastUsedRow * 2);
                        count++;
                    }

                    //close and release
                    xlWorkbook.Close();

                    //quit and release
                    xlApp.Quit();

                    UpdateStatus("Updating Database", count, lastUsedRow * 2);

                    foreach (CliObj cl in clients)
                    {
                        if (err.doesPTExist(cl.PT, _jurisUtility))
                        {
                            string sql = "select count(*) as CT, PTRFeeSch, PTRPrsTyp from PersTypRate " +
                                          " where PTRFeeSch = '" + clientRateSched + "' and PTRPrsTyp = '" + cl.PT + "' " +
                                          " group by PTRFeeSch, PTRPrsTyp " +
                                          " having count(*) > 0";

                            DataSet fd = _jurisUtility.RecordsetFromSQL(sql);
                            if (fd != null && fd.Tables.Count > 0 && fd.Tables[0].Rows.Count > 0)
                            {
                                string innersql = "update PersTypRate set PTRRate = cast(" + cl.rate + " as money) where PTRFeeSch = '" + clientRateSched + "' and PTRPrsTyp = '" + cl.PT + "' ";
                                _jurisUtility.ExecuteNonQuery(0, innersql);
                            }
                            else
                            {
                                string innersql = "insert into PersTypRate (PTRFeeSch, PTRPrsTyp, PTRRate) values ('" + clientRateSched + "', '" + cl.PT + "', cast(" + cl.rate + " as money))";
                                _jurisUtility.ExecuteNonQuery(0, innersql);

                            }
                        }
                        else
                        {
                            cl.error = true;
                            cl.errorMess = "Personnel Type " + cl.PT + " is not valid. Please update the PT for that record";
                        }
                        count++;
                        UpdateStatus("Updating Database", count, lastUsedRow * 2);

                    }
                    UpdateStatus("Updating Database", lastUsedRow * 2, lastUsedRow * 2);
                    count++;



                    UpdateStatus("All items updated.", 1, 1);

                    List<CliObj> errors = clients.Where(p => p.error == true).ToList();

                    if (errors.Count == 0)
                        MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                    else
                    {
                        DialogResult dr = MessageBox.Show("The process is complete but there were errors." + "\r\n" + "Would you like to see the error list?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.None);
                        if (dr == DialogResult.Yes)
                        {
                            DataSet df = new DataSet();
                            df.Tables.Add(err.ToDataTable(errors));
                            ReportDisplay rd = new ReportDisplay(df);
                            rd.ShowDialog();
                        }
                    }

                    clients.Clear();
                    fileName = "";
                }
                else
                    MessageBox.Show("That Client Code is not valid. Please update the spreadsheet", "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            fileName = "";
        }



        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private void buttonExcel_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Browse for Excel File";
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                System.Environment.Exit(0);

            }
            catch (Exception bbf)
            {
                Application.Exit();
            }
        }
    }
}
