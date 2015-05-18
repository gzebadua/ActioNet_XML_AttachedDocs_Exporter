using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Diagnostics;

namespace XML_Docs_Exporter
{
    public partial class IAAXMLExporter : Form
    {
        public IAAXMLExporter()
        {
            InitializeComponent();
        }

        public static string StringConnection()
        {
            return "Provider=SQLNCLI10;Server=Vdbw001tv\\msapps;Database=Staging;Trusted_Connection=Yes;";
        }

        public void execSQLCommand(string query)
        {
            OleDbConnection oleConnection = new OleDbConnection(StringConnection());
            OleDbCommand oleCommand = new OleDbCommand();

            oleConnection.Open();
            oleCommand.CommandText = query;
            oleCommand.Connection = oleConnection;
            oleCommand.ExecuteNonQuery();
            oleConnection.Close();
        }

        public string getXMLValue(string line, string field, string defaultValue = "")
        {

            int charsToCut = 0;

            try
            {
                if (line.Contains("<" + field + "></") == true) //empty
                {

                    return defaultValue;


                }
                else if (line.Contains("<" + field + ">") == true & line.Contains("<" + field + "></") == false) //Contains something
                { 

                    charsToCut = line.IndexOf("</" + field + ">") - line.IndexOf("<" + field + ">") - (field.Length + 2);

                    return line.Substring(line.IndexOf("<" + field + ">") + field.Length + 2, charsToCut);

                }

                return "";
            }
            catch (Exception ex)
            {
                string exceptionLine = line;
                int indexOf = exceptionLine.IndexOf("<" + field + ">");
                int fieldLength = field.Length + 2;
                int chars = charsToCut;
                return "";
            }

        }


        private void extractInfo()
        {

            btnExit.Enabled = false;
            Stopwatch timePerParse;
            string processRunningTime = "";

            try
            {
                inputFolder.ShowNewFolderButton = false;
                inputFolder.Description = "Please choose the folder where the outdated forms are";


                if (inputFolder.ShowDialog() == DialogResult.OK)
                {
                    timePerParse = Stopwatch.StartNew();

                    DirectoryInfo dirInfo = new DirectoryInfo(inputFolder.SelectedPath);

                    int fileCounter = 0;
                    foreach (FileInfo fi in dirInfo.GetFiles())
                    {

                        string filename = fi.Name;
                        fileCounter++;

                        rtb.AppendText("Process started at " + DateTime.Now + "\r\n");
                        //rtb.AppendText("Reading " + inputFolder.SelectedPath + "\\" + filename + "\r\n");

                        StreamReader reader = new StreamReader(inputFolder.SelectedPath + "\\" + filename);
                        string line = reader.ReadToEnd();
                        reader.Dispose();

                        string IAANumber = getXMLValue(line, "my:NewIAANo", ""); //NULL VALUES STAY AS BLANK IN DATABASE

                        string IAAdate = getXMLValue(line, "my:EstimatedCompletionDate", ""); 
                        string IAAdate2 = getXMLValue(line, "my:BudgetAnalystSignedDate", ""); 
                        string IAAdate3 = getXMLValue(line, "my:DivisionChiefSignedDate", ""); 
                        string IAAdate4 = getXMLValue(line, "my:AcquisitionSignedDate", ""); 
                        string IAAdate5 = getXMLValue(line, "my:DeputyAssociateAdminSignedDate", ""); 
                        string IAAdate6 = getXMLValue(line, "my:DirectorSignedDate", ""); 
                        string IAAdate7 = getXMLValue(line, "my:LegalSignedDate", ""); 
                        string IAAdate8 = getXMLValue(line, "my:CFOSignedDate", ""); 
                        string IAAdate9 = getXMLValue(line, "my:CreatedDate", ""); 

                        //string IAADocumentType = "FinalSOWPDF";
                        //int IAADocumentNumber = 1;
                        string IAAsowPDF = getXMLValue(line, "my:SignedFinalSOWPdf", "");
                        
                        //IAADocumentType = "FinalSOWWord";
                        //IAADocumentNumber = 2;
                        string IAAsowWord = getXMLValue(line, "my:SignedFinalSOWWord", "");
                        
                        //IAADocumentType = "FundingDocument";
                        //IAADocumentNumber = 3;
                        string IAAFundingDoc = getXMLValue(line, "my:FundingDocument", "");

                        //IAADocumentType = "AdditionalDocument";
                        //string IAAAdditionalDocs = "";
                        string line2 = line;
                        
                        List<string> additionalDocs = new List<string>();
                        int IAADocumentNumber = 0;

                        while (line2.IndexOf("<my:AdditionalDocument>") > -1)
                        {
                            IAADocumentNumber++;
                            additionalDocs.Add(getXMLValue(line2, "my:AdditionalDocument", ""));
                            
                            try
                            {
                                line2 = line2.Substring(line2.IndexOf("</my:AdditionalDocument>") + 24); //24 is the length of "</my:AdditionalDocument>"
                            }
                            catch (Exception ex1)
                            {
                                rtb.AppendText("\r\n\r\nEXCEPTION at " + inputFolder.SelectedPath + "\\" + filename + ": " + ex1.ToString() + "\r\n\r\n");
                            }
                            //IAAAdditionalDocs = "";
                        }

                        for (int i = IAADocumentNumber; i <= 6; i++)
                        {
                            additionalDocs.Add("");
                        }

                        execSQLCommand("INSERT INTO dbo.HistoricalIAADocuments VALUES('" + IAANumber + "', '" +
                        IAAdate + "', '" +
                        IAAdate2 + "', '" +
                        IAAdate3 + "', '" +
                        IAAdate4 + "', '" +
                        IAAdate5 + "', '" +
                        IAAdate6 + "', '" +
                        IAAdate7 + "', '" +
                        IAAdate8 + "', '" +
                        IAAdate9 + "', " +
                        "CONVERT(varbinary(max), '" + IAAsowPDF + "'), " +
                        "CONVERT(varbinary(max), '" + IAAsowWord + "'), " +
                        "CONVERT(varbinary(max), '" + IAAFundingDoc + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[0] + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[1] + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[2] + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[3] + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[4] + "'), " +
                        "CONVERT(varbinary(max), '" + additionalDocs[5] + "') " + 
                         ");");

                        filename = "";
                        line = "";
                        IAANumber = "";
                        IAAdate = "";
                        IAAsowPDF = "";
                        IAAsowWord = "";
                        IAAFundingDoc = "";
                        additionalDocs.Clear();
                        line2 = "";

                        //break;
                    }

                    timePerParse.Stop();
                    processRunningTime = timePerParse.ElapsedMilliseconds.ToString();
                }
                
            }
            catch (Exception ex)
            {
                rtb.AppendText("\r\n\r\nEXCEPTION :" + ex.ToString() + "\r\n\r\n");
            }

            rtb.AppendText("Process stopped at " + DateTime.Now + ". Duration : " + processRunningTime + " ms.\r\n");
            btnExit.Enabled = true;

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            extractInfo();
        }

    }
}
