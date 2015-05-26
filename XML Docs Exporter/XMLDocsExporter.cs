using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace XML_Docs_Exporter
{
	public partial class XMLDocsExporter : Form
	{

		public XMLDocsExporter()
		{
			InitializeComponent();
		}

		public static string StringConnection()
		{
			return "Server=Vdbw001tv\\msapps;Database=Staging;Trusted_Connection=True;";
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
            DateTime startTime = DateTime.Now;
            DateTime endTime = DateTime.Now;
            TimeSpan elapsedTime = endTime - startTime;

			try
			{
				inputFolder.ShowNewFolderButton = false;
				inputFolder.Description = "Please choose the folder where the XML InfoPath forms are stored";


				if (inputFolder.ShowDialog() == DialogResult.OK)
				{
                    startTime = DateTime.Now;

					DirectoryInfo dirInfo = new DirectoryInfo(inputFolder.SelectedPath);

					int fileCounter = 0;
                    rtb.AppendText("Process started at " + startTime + "\r\n");

					foreach (FileInfo fi in dirInfo.GetFiles())
					{

						const int SP1Header_Size = 20;
						const int FIXED_HEADER = 16;

						string filename = fi.Name;
						fileCounter++;
                        rtb.AppendText("Reading file #" + fileCounter + " : " + inputFolder.SelectedPath + "\\" + filename + "\r\n"); //Just a "receipt" of what was processed

						StreamReader reader = new StreamReader(inputFolder.SelectedPath + "\\" + filename);
						string line = reader.ReadToEnd();
						reader.Dispose();

						string IAANumber = getXMLValue(line, "my:NewIAANo", ""); 

						string IAAdate = getXMLValue(line, "my:EstimatedCompletionDate", ""); 
						string IAAdate2 = getXMLValue(line, "my:BudgetAnalystSignedDate", ""); 
						string IAAdate3 = getXMLValue(line, "my:DivisionChiefSignedDate", ""); 
						string IAAdate4 = getXMLValue(line, "my:AcquisitionSignedDate", ""); 
						string IAAdate5 = getXMLValue(line, "my:DeputyAssociateAdminSignedDate", ""); 
						string IAAdate6 = getXMLValue(line, "my:DirectorSignedDate", ""); 
						string IAAdate7 = getXMLValue(line, "my:LegalSignedDate", ""); 
						string IAAdate8 = getXMLValue(line, "my:CFOSignedDate", ""); 
						string IAAdate9 = getXMLValue(line, "my:CreatedDate", "");

						string IAAsowPDFAttachmentName = "";
						string IAAsowPDFEncodedString = getXMLValue(line, "my:SignedFinalSOWPdf", "");
                        byte[] IAAsowPDF = new byte[0];

						if (IAAsowPDFEncodedString.Length > 0) 
						{ 
							using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(IAAsowPDFEncodedString)))
							{
								int fileSize;
								int attachmentNameLength;
								BinaryReader theReader = new BinaryReader(ms);
								byte[] headerData = new byte[FIXED_HEADER];
								headerData = theReader.ReadBytes(headerData.Length);
								fileSize = (int)theReader.ReadUInt32();
								attachmentNameLength = (int)theReader.ReadUInt32() * 2;
								byte[] fileNameBytes = theReader.ReadBytes(attachmentNameLength);
								Encoding enc = Encoding.Unicode;
								IAAsowPDFAttachmentName = enc.GetString(fileNameBytes, 0, attachmentNameLength - 2);
								IAAsowPDF = theReader.ReadBytes(fileSize);
							}
						}
						string IAAsowWordAttachmentName = "";
						string IAAsowWordEncodedString = getXMLValue(line, "my:SignedFinalSOWWord", "");
                        byte[] IAAsowWord = new byte[0];

						if (IAAsowWordEncodedString.Length > 0)
						{
							using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(IAAsowWordEncodedString)))
							{
								int fileSize;
								int attachmentNameLength;
								BinaryReader theReader = new BinaryReader(ms);
								byte[] headerData = new byte[FIXED_HEADER];
								headerData = theReader.ReadBytes(headerData.Length);
								fileSize = (int)theReader.ReadUInt32();
								attachmentNameLength = (int)theReader.ReadUInt32() * 2;
								byte[] fileNameBytes = theReader.ReadBytes(attachmentNameLength);
								Encoding enc = Encoding.Unicode;
								IAAsowWordAttachmentName = enc.GetString(fileNameBytes, 0, attachmentNameLength - 2);
								IAAsowWord = theReader.ReadBytes(fileSize);
							}
						}
						string IAAFundingDocAttachmentName = "";
						string IAAFundingDocEncodedString = getXMLValue(line, "my:FundingDocument", "");
                        byte[] IAAFundingDoc = new byte[0];

						if (IAAFundingDocEncodedString.Length > 0)
						{
							using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(IAAFundingDocEncodedString)))
							{
								int fileSize;
								int attachmentNameLength;
								BinaryReader theReader = new BinaryReader(ms);
								byte[] headerData = new byte[FIXED_HEADER];
								headerData = theReader.ReadBytes(headerData.Length);
								fileSize = (int)theReader.ReadUInt32();
								attachmentNameLength = (int)theReader.ReadUInt32() * 2;
								byte[] fileNameBytes = theReader.ReadBytes(attachmentNameLength);
								Encoding enc = Encoding.Unicode;
								IAAFundingDocAttachmentName = enc.GetString(fileNameBytes, 0, attachmentNameLength - 2);
								IAAFundingDoc = theReader.ReadBytes(fileSize);
							}
						}

						string line2 = line;
						int IAADocumentNumber = 0;

						List<byte[]> additionalDocs = new List<byte[]>();
						List<string> additionalDocsAttachmentNames = new List<string>();

						string AdditionalDocumentAttachmentName = "";
						byte[] AdditionalDocument = new byte[0];

						while (line2.IndexOf("<my:AdditionalDocument>") > -1)
						{
							IAADocumentNumber++;

                            try
                            {
                                using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(getXMLValue(line2, "my:AdditionalDocument", ""))))
                                {
                                    int fileSize;
                                    int attachmentNameLength;
                                    BinaryReader theReader = new BinaryReader(ms);
                                    byte[] headerData = new byte[FIXED_HEADER];
                                    headerData = theReader.ReadBytes(headerData.Length);
                                    fileSize = (int)theReader.ReadUInt32();
                                    attachmentNameLength = (int)theReader.ReadUInt32() * 2;
                                    byte[] fileNameBytes = theReader.ReadBytes(attachmentNameLength);
                                    Encoding enc = Encoding.Unicode;
                                    AdditionalDocumentAttachmentName = enc.GetString(fileNameBytes, 0, attachmentNameLength - 2);
                                    AdditionalDocument = theReader.ReadBytes(fileSize);
                                }
                            
							    additionalDocs.Add(AdditionalDocument);
							    additionalDocsAttachmentNames.Add(AdditionalDocumentAttachmentName);

							    AdditionalDocument = new byte[0];
							    AdditionalDocumentAttachmentName = "";
							
							    try
							    {
								    line2 = line2.Substring(line2.IndexOf("</my:AdditionalDocument>") + 24); //24 is the length of "</my:AdditionalDocument>"
							    }
							    catch (Exception ex1)
							    {
								    rtb.AppendText("\r\n\r\nEXCEPTION at " + inputFolder.SelectedPath + "\\" + filename + ": " + ex1.ToString() + "\r\n\r\n");
							    }
                            }
                            catch (Exception ex)
                            {

                                rtb.AppendText("\r\n\r\nCorruption on file # " + IAADocumentNumber + ". Skipping to the next additional document...\r\n\r\n");
                                
                                AdditionalDocument = new byte[0];
                                AdditionalDocumentAttachmentName = "";
                                IAADocumentNumber--;

                                try
                                {
                                    line2 = line2.Substring(line2.IndexOf("</my:AdditionalDocument>") + 24); //24 is the length of "</my:AdditionalDocument>"
                                }
                                catch (Exception ex1)
                                {
                                    rtb.AppendText("\r\n\r\nEXCEPTION at " + inputFolder.SelectedPath + "\\" + filename + ": " + ex1.ToString() + "\r\n\r\n");
                                }

                            }
						}

                        string strCommand =
                        "INSERT INTO dbo.HistoricalIAADocuments VALUES('" +
                        IAANumber + "', '" +
                        IAAdate + "', '" +
                        IAAdate2 + "', '" +
                        IAAdate3 + "', '" +
                        IAAdate4 + "', '" +
                        IAAdate5 + "', '" +
                        IAAdate6 + "', '" +
                        IAAdate7 + "', '" +
                        IAAdate8 + "', '" +
                        IAAdate9 + "', ";

                        if (IAAsowPDFEncodedString.Length > 0)
                        {
                            strCommand += "@IAAsowPDF, '" +
                            IAAsowPDFAttachmentName.Replace("'", "") + "', ";
                        }
                        else
                        {
                            strCommand += "NULL, NULL, ";
                        }

                        if (IAAsowWordEncodedString.Length > 0)
                        {
                            strCommand += "@IAAsowWord, '" +
                            IAAsowWordAttachmentName.Replace("'", "") + "', ";
                        }
                        else
                        {
                            strCommand += "NULL, NULL, ";
                        }

                        if (IAAFundingDocEncodedString.Length > 0)
                        {
                            strCommand += "@IAAFundingDoc, '" +
                            IAAFundingDocAttachmentName.Replace("'", "") + "', ";
                        }
                        else
                        {
                            strCommand += "NULL, NULL, ";
                        }
						
						if (IAADocumentNumber > 0) { 
							for (int i = 1; i <= IAADocumentNumber; i++)
							{
								strCommand += "@additionalDocs" + (i) + ", '" +
                                additionalDocsAttachmentNames[i - 1].Replace("'", "") + "', ";
							}
						}

                        //if (IAADocumentNumber > 7) 
                        //{
                        //    rtb.AppendText("\r\n");
                        //    rtb.AppendText("\r\n");
                        //    rtb.AppendText("\r\n");
                        //    rtb.AppendText("There are " + IAADocumentNumber + " AdditionalDocuments in this XML file: " + filename + " ! \r\n");
                        //    rtb.AppendText("\r\n");
                        //    rtb.AppendText("\r\n");
                        //    rtb.AppendText("\r\n");
                        //}

						if (IAADocumentNumber < 7)
						{
							for (int i = IAADocumentNumber; i+1 <= 7; i++)
							{
								strCommand += "NULL, NULL";
								if (i+1 < 7)
								{
									strCommand += ", ";
								}
								else if (i+1 == 7)
								{
									strCommand += ");";
								}
							}

						}
						else if (IAADocumentNumber == 7)
						{
                            strCommand = strCommand.Substring(0, strCommand.Length - 2); // Removing the ", " since the pattern is full
							strCommand += ");";
						}

						using (SqlConnection dbConnection = new SqlConnection(StringConnection())) { 

							using (SqlCommand cmd = new SqlCommand(strCommand, dbConnection))
							{

                                if (IAAsowPDFEncodedString.Length > 0) { cmd.Parameters.Add("@IAAsowPDF", SqlDbType.VarBinary, -1).Value = IAAsowPDF; }
                                if (IAAsowWordEncodedString.Length > 0) { cmd.Parameters.Add("@IAAsowWord", SqlDbType.VarBinary, -1).Value = IAAsowWord; }
                                if (IAAFundingDocEncodedString.Length > 0) { cmd.Parameters.Add("@IAAFundingDoc", SqlDbType.VarBinary, -1).Value = IAAFundingDoc; }

								if (IAADocumentNumber > 0)
								{
									for (int i = 1; i <= IAADocumentNumber; i++)
									{
										cmd.Parameters.Add("@additionalDocs" + (i), SqlDbType.VarBinary, -1).Value = additionalDocs[(i-1)];
									}
								}
								dbConnection.Open();
								cmd.ExecuteNonQuery();
								dbConnection.Close();
							}
						}

						filename = "";
						line = "";
						IAANumber = "";
						IAAdate = "";
						IAAdate2 = "";
						IAAdate3 = "";
						IAAdate4 = "";
						IAAdate5 = "";
						IAAdate6 = "";
						IAAdate7 = "";
						IAAdate8 = "";
						IAAdate9 = "";
                        IAAsowPDF = new byte[0];
						IAAsowPDFAttachmentName = "";
						IAAsowPDFEncodedString = "";
						IAAsowWord = new byte[0];
						IAAsowWordAttachmentName = "";
						IAAsowWordEncodedString = "";
                        IAAFundingDoc = new byte[0];
						IAAFundingDocAttachmentName = "";
						IAAFundingDocEncodedString = "";
						additionalDocs.Clear();
						additionalDocsAttachmentNames.Clear();
                        AdditionalDocument = new byte[0];
                        AdditionalDocumentAttachmentName = "";
						line2 = "";
						strCommand = "";
						IAADocumentNumber = 0;
						
                        //break; //just for debugging purposes

                        fi.Delete(); //Make a copy of the XML files first as a backup. This deletes the processed files so you end up with only problematic files
					}

                    endTime = DateTime.Now;
                    elapsedTime = endTime - startTime;
				}
				
			}
			catch (Exception ex)
			{
				rtb.AppendText("\r\n\r\nEXCEPTION :" + ex.ToString() + "\r\n\r\n");
			}

            rtb.AppendText("Process stopped at " + endTime.ToLocalTime() + ". Duration : " + elapsedTime.TotalMinutes + " ms.\r\n");
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

        private void rtb_TextChanged(object sender, EventArgs e)
        {
            rtb.SelectionStart = rtb.Text.Length; //Set the current caret position at the end
            rtb.ScrollToCaret(); //Now scroll it automatically
        }

	}
}
