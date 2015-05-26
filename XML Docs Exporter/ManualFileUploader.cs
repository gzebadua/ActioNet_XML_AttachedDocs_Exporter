using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace XML_Docs_Exporter
{
    public partial class ManualFileUploader : Form
    {
        public ManualFileUploader()
        {
            InitializeComponent();
        }

        public static string StringConnection()
        {
            return "Server=Vdbw001tv\\msapps;Database=Staging;Trusted_Connection=True;";
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {

            try
            {

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string strFileName = openFileDialog.SafeFileName;
                    byte[] IAAFile = new byte[0];

                    IAAFile = File.ReadAllBytes(openFileDialog.FileName);

                    string strCommand =
                    "UPDATE    dbo.HistoricalIAADocuments SET " +
                    "AdditionalDocument3FileName = '" + strFileName + "', " +
                    "AdditionalDocument3 = @additionalDocument3 " +
                    "WHERE     (IAANumber = 'FA63C1') AND (FinalSOWPDFFileName = 'DTFAVP-13-X-00091 (FA63C1).pdf')";

                    using (SqlConnection dbConnection = new SqlConnection(StringConnection()))
                    {

                        using (SqlCommand cmd = new SqlCommand(strCommand, dbConnection))
                        {

                            cmd.Parameters.Add("@additionalDocument3", SqlDbType.VarBinary, -1).Value = IAAFile;

                            dbConnection.Open();
                            cmd.ExecuteNonQuery();
                            dbConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
