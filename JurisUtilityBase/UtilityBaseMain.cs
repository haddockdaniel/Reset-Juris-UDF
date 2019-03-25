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

        private int totalNumberOfRecords = 0;

        public List<UDF> definedFields = new List<UDF>();

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

            uncheckAllCBsAndMakeThemInvisible();
            definedFields.Clear();
            totalNumberOfRecords = 0;

            String SQL = "select spname, sptxtvalue, replace(replace(spname, 'FLD',''), 'UDF#','  ') as UDF, " +
                " replace(left(sptxtvalue, charindex(',', sptxtvalue) -1),' ','') as UDFName , " +
                " replace(substring(sptxtvalue, charindex(',', sptxtvalue), charindex(',', sptxtvalue, charindex(',', sptxtvalue) + 1) - charindex(',', sptxtvalue)),',','') as UDFType, " +
                " replace(substring(sptxtvalue,charindex(',', sptxtvalue, charindex(',', sptxtvalue) + 1), " +
                " charindex(',', sptxtvalue, charindex(',', sptxtvalue, charindex(',', sptxtvalue) + 1) +1) - " +
                " charindex(',', sptxtvalue, charindex(',', sptxtvalue) + 1)),',','') as UFDLen, right(sptxtvalue,1) as UDFReq " +
                " from sysparam where spname like '%UDF%' and sptxtvalue not like '%,X,0,N' order by spname";

            DataSet myRSFS = _jurisUtility.RecordsetFromSQL(SQL);

            UDF udf = null;

            if (myRSFS.Tables[0].Rows.Count == 0)
                MessageBox.Show("This Juris Database has no UDF fields defined", "No UDF Fields", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            else
            {
                foreach (DataRow dr in myRSFS.Tables[0].Rows)
                {
                    udf = new UDF();
                    udf.spName = dr["spname"].ToString();
                    udf.UDFname = dr["UDFName"].ToString();
                    definedFields.Add(udf);
                    totalNumberOfRecords++;
                }
            }

            showCheckBoxesThatAreInDB();

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            int count = 1;
            foreach (Control ctrl in this.Controls)
            {
                // You can use the following if condition to target the specific control
                // if (ctrl.Name.Equals("groupBox1"))
                string SQL = "";
                if (ctrl.ToString().StartsWith("System.Windows.Forms.GroupBox"))
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c is CheckBox && ((CheckBox)c).Checked && ((CheckBox)c).Visible)
                        {
                            string text = ((CheckBox)c).Text;
                            string txtValBeg = text.Replace("Fld", "").Substring(0,1);
                            string txtValEnd = text.Substring(text.Length - 5, 5);
                            txtValEnd = txtValEnd.Replace("#", "");
                            txtValEnd = txtValEnd.ToUpper();
                            string table = text.Replace("Fld", "");
                            table = table.Replace("UDF#", "");
                            table = table.Substring(0, table.Length - 1);

                            SQL = "update sysparam " +
                                "set sptxtvalue='" + txtValBeg + " " + txtValEnd + ",X,0,N' " +
                                "where spname='" + text + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            UDF currentUDF = definedFields.First(s => text == s.spName);
                            string column = currentUDF.UDFname;

                            SQL = "ALTER TABLE " + table + " NOCHECK CONSTRAINT ALL";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            SQL = "ALTER TABLE " + table + " drop column " + column;
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            SQL = "ALTER TABLE " + table + " CHECK CONSTRAINT ALL";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            UpdateStatus("Resetting field: " + text, count, totalNumberOfRecords + 1);
                            count++;
                        }
                    }
                }
            }

            UpdateStatus("All UDF fields reset.", totalNumberOfRecords + 1, totalNumberOfRecords + 1);

            MessageBox.Show("The process is complete", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
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

        private void uncheckAllCBsAndMakeThemInvisible()
        {
            foreach (Control ctrl in this.Controls)
            {
                // You can use the following if condition to target the specific control
                // if (ctrl.Name.Equals("groupBox1"))
                if (ctrl.ToString().StartsWith("System.Windows.Forms.GroupBox"))
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c is CheckBox)
                        {
                            ((CheckBox)c).Checked = false;
                            ((CheckBox)c).Visible = false;
                        }
                    }
                }
            }

        }

        private void showCheckBoxesThatAreInDB()
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl.ToString().StartsWith("System.Windows.Forms.GroupBox"))
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c is CheckBox)
                        {
                            foreach (UDF field in definedFields)
                            {
                                if (((CheckBox)c).Text == field.spName)
                                    ((CheckBox)c).Visible = true;
                            }
                        }
                    }
                }
            }

        }

        private void labelDescription_Click(object sender, EventArgs e)
        {

        }


    }
}
