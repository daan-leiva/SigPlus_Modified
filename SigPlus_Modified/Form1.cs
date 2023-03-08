using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using System.IO;
/*
 * Coordinate system: 0 - Logical Tablet Coordinates, 1 - LCD Coordinates
 * Pad is 4x3 : 240, 1280 (LCD Coordinates)
 * Destination: 0 - foreground,  - background
 * 
 * Hot Spot List:
 * 0 - Opening screen tap anywhere
 * 1 - clear button for name screen
 * 2 - ok button for name screen
 * 3 - clear button for company name screen
 * 4 - ok button for company name screen
 * 5 - yes citizen
 * 6 - no citizen
 */


namespace SigPlus_Modified
{
    public partial class Form1 : Form
    {
        Bitmap mainScreen, nameScreen, companyNameScreen, usCitizenCheckboxes, check, atiLogo, sigField, peopleCheckboxes;
        Topaz.SigPlusNET sigPlusNET1 = new Topaz.SigPlusNET();
        DataRow databaseValues = new DataRow();
        Font textFont = new System.Drawing.Font("Arial", 11.0F, System.Drawing.FontStyle.Regular);
        Font textFontOther = new System.Drawing.Font("Arial", 11.0F, System.Drawing.FontStyle.Italic);
        List<viewRow> bindingSource = new List<viewRow>();
        string fileChosen = String.Empty;
        // OpenFileDialog dialog = new OpenFileDialog();
        string[] visiteeNames = { "Brittany B.", "Dwight H.", "Evan H.", "Kirk W.", "Kyle S.", "Nara P.", "Robin A.", "Sheldon B.", "Tim B.", "Tom D.", "Tony C.", "Other" };
        string username = "jbread";
        string password = "Cloudy2Day";
        string server = "ATI-SQL";
        string database = "ATI_SignIn";

        public Form1()
        {
            InitializeComponent();

            sigPlusNET1.PenUp += new System.EventHandler(this.sigPlusNET1_PenUp);
            this.Controls.Add(this.sigPlusNET1);
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold);
            textFont = new Font(textFont, FontStyle.Bold);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //
            // FILE SET UP
            //
            /*
            dialog.InitialDirectory = "c:\\";
            dialog.Filter = "Database Files (*.accdb)|*.accdb";
            dialog.FilterIndex = 2;
            dialog.RestoreDirectory = true;
            dialog.Multiselect = false;
            bool fileSelected = false;

            while (!fileSelected)
            {
               if (dialog.ShowDialog() == DialogResult.OK)
               {
                  try
                  {
                     if (dialog.CheckFileExists)
                     {
                        fileChosen = dialog.FileName;
                        fileSelected = true;
                     }
                  }
                  catch (Exception ex)
                  {
                     MessageBox.Show("Error: Could not read file from disk. Contact IT support. Original error: " + ex.Message);
                     Application.Exit();
                     Environment.Exit(1);
                  }
               }
               else
               {
                     Application.Exit();
                     Environment.Exit(1);
               }
            }
            */
            //
            // SET UP DATABASE GRIDVIEW
            //
            LoadDatabase();

            sigPlusNET1.SetTabletState(1); //open port, turn tablet on
            sigPlusNET1.ClearTablet(); //Clears the SigPlus object of ink

            sigPlusNET1.SetLCDCaptureMode(2);
            sigPlusNET1.LCDSetWindow(0, 0, 1, 1); //Prohibit inking on entire LCD
            sigPlusNET1.SetSigWindow(1, 0, 0, 1, 1); //Prohibit inking in SigPlus
            sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128); //Refresh entire tablet
            sigPlusNET1.SetTranslateBitmapEnable(false);

            //
            // SET UP HOT SPOTS
            // Using LCD Coordinates
            //
            sigPlusNET1.KeyPadAddHotSpot(0, 1, 0, 0, 240, 128); // Hot spot for first screen "tap anywhere on screen to start"

            //
            // Load images to use
            //
            mainScreen = new System.Drawing.Bitmap(Application.StartupPath + "\\prescription.bmp");
            nameScreen = new System.Drawing.Bitmap(Application.StartupPath + "\\screen2.bmp");
            companyNameScreen = new System.Drawing.Bitmap(Application.StartupPath + "\\screen2.bmp");
            usCitizenCheckboxes = new System.Drawing.Bitmap(Application.StartupPath + "\\citizenQ.bmp");
            check = new System.Drawing.Bitmap(Application.StartupPath + "\\check.bmp");
            atiLogo = new System.Drawing.Bitmap(Application.StartupPath + "\\logo.bmp");
            sigField = new System.Drawing.Bitmap(Application.StartupPath + "\\signatureBox2.bmp");
            peopleCheckboxes = new System.Drawing.Bitmap(Application.StartupPath + "\\peopleQ.bmp");

            //
            // Load foreground image
            //
            sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "            Welcome to ATI\n       Tap anywhere to begin");
            sigPlusNET1.LCDSendGraphic(0, 2, 58, 45, atiLogo);

            sigPlusNET1.SetLCDCaptureMode(2);

            this.BringToFront();
        }

        private void LoadDatabase()
        {
            bindingSource.Clear();

            Image nameImage;
            Image companyImage;
            Image otherPersonImage;

            using (OleDbConnection connection = new OleDbConnection("Provider=sqloledb;Data Source=" + this.server + ";Initial Catalog=" + this.database + ";User Id=" + this.username + ";Password=" + this.password + ";"))
            using (OleDbCommand command = new OleDbCommand("SELECT '', TimeStamp , NameByteString , CompanyRepresentedByteStream , CitizenOrResident ,Visitee ,OtherPerson , Chaperone FROM ATI_SignIn.dbo.VisitorsTable ORDER BY TimeStamp;", connection))
            {
                connection.Open();
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    sigPlusNET1.SetImageXSize(300);
                    sigPlusNET1.SetImageYSize(120);

                    while (reader.Read())
                    {
                        //MessageBox.Show(string.Format(reader.GetDateTime(0).ToString(), "Testtttt", MessageBoxButtons.OK, MessageBoxIcon.Error));
                        // set the name string back into the sigplus
                        sigPlusNET1.SetSigString(reader.IsDBNull(1) ? string.Empty : reader.GetString(2));
                        nameImage = sigPlusNET1.GetSigImage();
                        sigPlusNET1.SetSigString(reader.IsDBNull(3) ? string.Empty : reader.GetString(3));
                        companyImage = sigPlusNET1.GetSigImage();
                        sigPlusNET1.SetSigString(reader.IsDBNull(6) ? string.Empty : reader.GetString(6));
                        otherPersonImage = sigPlusNET1.GetSigImage();
                        bindingSource.Add(new viewRow(
                           reader.GetDateTime(1),
                           nameImage, companyImage,
                           reader.GetBoolean(4),
                           reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                           otherPersonImage,
                           reader.IsDBNull(7) ? string.Empty : reader.GetString(7)));
                    }

                    startDateTimePicker.ValueChanged -= this.startDateTimePicker_ValueChanged;
                    if (bindingSource.Count != 0)
                    {
                        startDateTimePicker.MaxDate = bindingSource.ElementAt(bindingSource.Count - 1).timeStamp;
                        startDateTimePicker.MinDate = bindingSource.ElementAt(0).timeStamp;
                    }
                    else
                    {
                        startDateTimePicker.MinDate = DateTime.Now;
                        startDateTimePicker.MaxDate = DateTime.Now;
                    }
                    startDateTimePicker.Value = startDateTimePicker.MinDate;
                    startDateTimePicker.ValueChanged += this.startDateTimePicker_ValueChanged;

                    endDateTimePicker.ValueChanged -= this.endDateTimePicker_ValueChanged;
                    if (bindingSource.Count != 0)
                    {
                        endDateTimePicker.MaxDate = bindingSource.ElementAt(bindingSource.Count - 1).timeStamp;
                        endDateTimePicker.MinDate = bindingSource.ElementAt(0).timeStamp;
                    }
                    else
                    {
                        endDateTimePicker.MinDate = DateTime.Now;
                        endDateTimePicker.MaxDate = DateTime.Now;
                    }
                    endDateTimePicker.Value = endDateTimePicker.MaxDate;
                    endDateTimePicker.ValueChanged += this.endDateTimePicker_ValueChanged;

                    dataGridView1.DataSource = bindingSource;
                    dataGridView1.Columns[0].HeaderCell.Value = "Time Stamp";
                    dataGridView1.Columns[1].HeaderCell.Value = "Name";
                    dataGridView1.Columns[2].HeaderCell.Value = "Company";
                    dataGridView1.Columns[3].HeaderCell.Value = "Citizen?";
                    dataGridView1.Columns[3].Width = 80;
                    dataGridView1.Columns[4].HeaderCell.Value = "Visitee";
                    dataGridView1.Columns[4].Width = (int)(dataGridView1.Columns[4].Width * 1.2);
                    dataGridView1.Columns[5].HeaderCell.Value = "Other";
                    dataGridView1.Columns[6].HeaderCell.Value = "Chaperone";

                    dataGridView1.Update();
                    dataGridView1.Refresh();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Database transfer error. Contact IT support.\n\n{0}", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error));
                    Application.Exit();
                    Environment.Exit(1);
                }
                connection.Close();
            }
        }

        private void ReloadDatabase()
        {
            bindingSource.Clear();

            Image nameImage;
            Image companyImage;
            Image otherPersonImage;

            using (OleDbConnection connection = new OleDbConnection("Provider=sqloledb;Data Source=" + this.server + ";Initial Catalog=" + this.database + ";User Id=" + this.username + ";Password=" + this.password + ";"))
            using (OleDbCommand command = new OleDbCommand("SELECT '', TimeStamp , NameByteString , CompanyRepresentedByteStream , CitizenOrResident ,Visitee ,OtherPerson , Chaperone FROM ATI_SignIn.dbo.VisitorsTable ORDER BY TimeStamp;", connection))
            {
                connection.Open();
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    sigPlusNET1.SetImageXSize(300);
                    sigPlusNET1.SetImageYSize(120);

                    while (reader.Read())
                    {
                        if (reader.GetDateTime(1).Date >= startDateTimePicker.Value.Date && reader.GetDateTime(1).Date <= endDateTimePicker.Value.Date)
                        {
                            // set the name string back into the sigplus
                            sigPlusNET1.SetSigString(reader.IsDBNull(2) ? string.Empty : reader.GetString(2));
                            nameImage = sigPlusNET1.GetSigImage();
                            sigPlusNET1.SetSigString(reader.IsDBNull(3) ? string.Empty : reader.GetString(3));
                            companyImage = sigPlusNET1.GetSigImage();
                            sigPlusNET1.SetSigString(reader.IsDBNull(6) ? string.Empty : reader.GetString(6));
                            otherPersonImage = sigPlusNET1.GetSigImage();
                            bindingSource.Add(new viewRow(
                               reader.GetDateTime(1),
                               nameImage, companyImage,
                               reader.GetBoolean(4),
                               reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                               otherPersonImage,
                               reader.IsDBNull(7) ? string.Empty : reader.GetString(7)));
                        }
                    }

                    dataGridView1.DataSource = null;
                    dataGridView1.DataSource = bindingSource;
                    dataGridView1.Columns[0].HeaderCell.Value = "Time Stamp";
                    dataGridView1.Columns[1].HeaderCell.Value = "Name";
                    dataGridView1.Columns[2].HeaderCell.Value = "Company";
                    dataGridView1.Columns[3].HeaderCell.Value = "Citizen?";
                    dataGridView1.Columns[3].Width = 70;
                    dataGridView1.Columns[4].HeaderCell.Value = "Visitee";
                    dataGridView1.Columns[4].Width = (int)(dataGridView1.Columns[4].Width * 1.2);
                    dataGridView1.Columns[5].HeaderCell.Value = "Other";

                    dataGridView1.Update();
                    dataGridView1.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Database transfer error. Contact IT support.\n\n{0}", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error));
                    Application.Exit();
                    Environment.Exit(1);
                }
                connection.Close();
            }
        }

        private void sigPlusNET1_PenUp(object sender, EventArgs e)
        {
            if (sigPlusNET1.KeyPadQueryHotSpot(0) > 0) // hitting intro screen
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
                sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Full Name:");
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                sigPlusNET1.KeyPadClearHotSpotList();

                sigPlusNET1.KeyPadAddHotSpot(1, 1, 106, 51, 36, 16); // Clear Button
                sigPlusNET1.KeyPadAddHotSpot(2, 1, 187, 52, 30, 15); // OK Button

                sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                          //to display ink
                sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                //specifies area in sigplus object to accept ink
                sigPlusNET1.SetLCDCaptureMode(2);
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(1) > 0) // clear button
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDRefresh(1, 106, 52, 40, 16); // invert clear button to show that use has pressed it
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                          //to display ink
                sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                //specifies area in sigplus object to accept ink
                sigPlusNET1.SetLCDCaptureMode(2);
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(2) > 0) // OK button
            {
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(1, 187, 53, 34, 15);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                // check that it's more than a point
                if (sigPlusNET1.NumberOfTabletPoints() > 0)
                {
                    databaseValues.name = sigPlusNET1.GetSigString(); //strSig now holds signature

                    //
                    // Set up company name screen
                    //
                    sigPlusNET1.ClearSigWindow(1);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Company Name:");
                    sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                    sigPlusNET1.KeyPadClearHotSpotList(); //clear out page 1 hotspots

                    //Add page 2 hotspots
                    sigPlusNET1.KeyPadAddHotSpot(3, 1, 106, 51, 36, 16); // Clear Button
                    sigPlusNET1.KeyPadAddHotSpot(4, 1, 187, 52, 30, 15); // OK Button

                    sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                              //to display ink
                    sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                    //specifies area in sigplus object to accept ink
                    sigPlusNET1.SetLCDCaptureMode(2);
                }
                else
                {
                    Font please = new System.Drawing.Font("Arial", 16.0F, System.Drawing.FontStyle.Regular);
                    sigPlusNET1.LCDWriteString(0, 2, 55, 38, please, "Please Sign");
                    sigPlusNET1.LCDWriteString(0, 2, 30, 63, please, "Before Continuing...");
                    System.Threading.Thread.Sleep(2500);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Full Name:");
                    sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                    sigPlusNET1.SetLCDCaptureMode(2);
                }
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(3) > 0) // clear button
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDRefresh(1, 106, 52, 40, 16); // invert clear button to show that use has pressed it
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                          //to display ink
                sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                //specifies area in sigplus object to accept ink
                sigPlusNET1.SetLCDCaptureMode(2);
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(4) > 0) // OK button
            {
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(1, 187, 53, 34, 15);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                // check that it's more than a point
                if (sigPlusNET1.NumberOfTabletPoints() > 0)
                {
                    databaseValues.companyName = sigPlusNET1.GetSigString(); //strSig now holds signature

                    //
                    // Set up citizen screen
                    //
                    sigPlusNET1.ClearSigWindow(1);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Are you a U.S. citizen or\na lawful permanent resident?");
                    sigPlusNET1.LCDSendGraphic(0, 2, 10, 64, usCitizenCheckboxes); // set name image
                    sigPlusNET1.LCDWriteString(0, 2, 80, 100, textFontOther, "ITAR 22 CFR 120.15");
                    sigPlusNET1.KeyPadClearHotSpotList(); //clear out page 1 hotspots
                                                          //Add page 2 hotspots

                    sigPlusNET1.KeyPadAddHotSpot(5, 1, 11, 55, 15, 25); // yes checkbox
                    sigPlusNET1.KeyPadAddHotSpot(6, 1, 11, 75, 15, 25); // no checkbox

                    sigPlusNET1.SetLCDCaptureMode(2);
                }
                else
                {
                    Font please = new System.Drawing.Font("Arial", 16.0F, System.Drawing.FontStyle.Regular);
                    sigPlusNET1.LCDWriteString(0, 2, 55, 38, please, "Please Sign");
                    sigPlusNET1.LCDWriteString(0, 2, 30, 63, please, "Before Continuing...");
                    System.Threading.Thread.Sleep(2500);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Company Name:");
                    sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                    sigPlusNET1.SetLCDCaptureMode(2);
                }
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(5) > 0) // yes checkbox for citizen
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDSendGraphic(0, 3, 14, 65, check);

                // sleep for a second so user can see that checkbox was checked off
                Thread.Sleep(1000);

                // Clear window
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                databaseValues.citizenOrResident = true;

                //
                // SET UP FOR RESIDENT SCREEN
                //
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.ClearTablet();
                sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Who are you visiting?");
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 40, peopleCheckboxes); // set checkbox image
                sigPlusNET1.KeyPadClearHotSpotList(); //clear out page 1 hotspots

                //Add new hotspots
                for (short i = 0; i < 3; i++) // x sweep
                    for (short j = 0; j < 4; j++)
                        sigPlusNET1.KeyPadAddHotSpot((short)(7 + i * 4 + j), 1, (short)(80 * i), (short)(40 + j * 23), 20, 15);

                sigPlusNET1.SetLCDCaptureMode(2);
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(6) > 0) // no checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                sigPlusNET1.LCDSendGraphic(0, 3, 14, 90, check);
                sigPlusNET1.SetLCDCaptureMode(2);

                // sleep for a second so user can see that checkbox was checked off
                Thread.Sleep(1000);

                // Clear window
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                databaseValues.citizenOrResident = false;

                //
                // SET UP FOR RESIDENT SCREEN
                //
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.ClearTablet();
                sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Who are you visiting?");
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 40, peopleCheckboxes); // set checkbox image
                sigPlusNET1.KeyPadClearHotSpotList(); //clear out page 1 hotspots

                //Add new hotspots
                for (short i = 0; i < 3; i++) // x sweep
                    for (short j = 0; j < 4; j++) // y sweep
                        sigPlusNET1.KeyPadAddHotSpot((short)(7 + i * 4 + j), 1, (short)(80 * i), (short)(40 + j * 23), 20, 15);

                sigPlusNET1.SetLCDCaptureMode(2);
            }
            #region peopleOptions
            else if (sigPlusNET1.KeyPadQueryHotSpot(7) > 0) // person1 checked
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[0];
                sigPlusNET1.LCDSendGraphic(0, 3, 4, 40 + 2, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(8) > 0) // person2 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[1];
                sigPlusNET1.LCDSendGraphic(0, 3, 4, 40 + 2 + 1 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(9) > 0) // person3 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[2];
                sigPlusNET1.LCDSendGraphic(0, 3, 4, 40 + 2 + 2 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(10) > 0) // person4 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[3];
                sigPlusNET1.LCDSendGraphic(0, 3, 4, 40 + 2 + 3 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(11) > 0) // person5 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = string.Empty;

                databaseValues.visitee = visiteeNames[0];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 1 * 80, 40 + 2 + 0 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(12) > 0) // person6 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[5];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 1 * 80, 40 + 2 + 1 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(13) > 0) // person7 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[6];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 1 * 80, 40 + 2 + 2 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(14) > 0) // person8 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[7];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 1 * 80, 40 + 2 + 3 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(15) > 0) // person9 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[8];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 2 * 80, 40 + 2 + 0 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(16) > 0) // person10 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[9];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 2 * 80, 40 + 2 + 1 * 23, check);

                repetitiveMethod();
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(17) > 0) // person11 checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[10];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 2 * 80, 40 + 2 + 2 * 23, check);

                repetitiveMethod();

            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(18) > 0) // other checkbox
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus

                //
                // find person/ mark correct box
                //
                databaseValues.visitee = visiteeNames[11];
                sigPlusNET1.LCDSendGraphic(0, 3, 0 + 4 + 2 * 80, 40 + 2 + 3 * 23, check);

                sigPlusNET1.SetLCDCaptureMode(2);

                // sleep for a second so user can see that checkbox was checked off
                Thread.Sleep(1000);

                // Clear window
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
                sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Full Name:");
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                sigPlusNET1.KeyPadClearHotSpotList();

                sigPlusNET1.KeyPadAddHotSpot(19, 1, 106, 51, 36, 16); // Clear Button
                sigPlusNET1.KeyPadAddHotSpot(20, 1, 187, 52, 30, 15); // OK Button

                sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                          //to display ink
                sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                //specifies area in sigplus object to accept ink
                sigPlusNET1.SetLCDCaptureMode(2);
            }
            #endregion
            else if (sigPlusNET1.KeyPadQueryHotSpot(19) > 0) // clear button for other
            {
                sigPlusNET1.ClearSigWindow(1); //clear hot spot buffer
                sigPlusNET1.ClearTablet(); //Clear SigPlus
                sigPlusNET1.LCDRefresh(1, 106, 52, 40, 16); // invert clear button to show that use has pressed it
                sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                sigPlusNET1.LCDSetWindow(2, 74, 236, 50); //Permits only the section on LCD
                                                          //to display ink
                sigPlusNET1.SetSigWindow(1, 0, 70, 240, 58);
                //specifies area in sigplus object to accept ink
                sigPlusNET1.SetLCDCaptureMode(2);
            }
            else if (sigPlusNET1.KeyPadQueryHotSpot(20) > 0) // ok button for other
            {
                sigPlusNET1.ClearSigWindow(1);
                sigPlusNET1.LCDRefresh(1, 187, 53, 34, 15);
                sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                // check that it's more than a point
                if (sigPlusNET1.NumberOfTabletPoints() > 0)
                {
                    databaseValues.otherPerson = sigPlusNET1.GetSigString(); //strSig now holds signature

                    databaseValues.chaperone = string.Empty;
                    // if not a citizen, request a chaperone
                    if (!databaseValues.citizenOrResident)
                        databaseValues.chaperone = Microsoft.VisualBasic.Interaction.InputBox("Visitor is not a citizen. Please assign a chaperone.", "Chaperone Text", "Enter chaperone's name here...");

                    databaseValues.timeStamp = DateTime.Now;

                    // Clear window
                    sigPlusNET1.ClearSigWindow(1);
                    sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);

                    Font thankyou = new System.Drawing.Font("Arial", 16.0F, System.Drawing.FontStyle.Regular);
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, thankyou, "Thank You For Signing");
                    sigPlusNET1.LCDSendGraphic(0, 2, 58, 45, atiLogo);
                    sigPlusNET1.SetTabletState(0); //turn off tablet to use justification below
                    int oldJustifyMode = sigPlusNET1.GetJustifyMode();
                    sigPlusNET1.SetJustifyMode(5); //this will zoom signature & justify to center
                    Thread.Sleep(500);

                    // CUT FROM HERE
                    using (OleDbConnection connection = new OleDbConnection("Provider=sqloledb;Data Source=" + this.server + ";Initial Catalog=" + this.database + ";User Id=" + this.username + ";Password=" + this.password + ";"))
                    using (OleDbCommand command = new OleDbCommand("INSERT INTO ATI_SignIn.dbo.VisitorsTable ([TimeStamp], [NameByteString], [CompanyRepresentedByteStream], [CitizenOrResident], [Visitee], [Chaperone]) VALUES (\'" + databaseValues.timeStamp.ToString() + "\', \'" + databaseValues.name + "\', \'" + databaseValues.companyName + "\', \'" + databaseValues.citizenOrResident + "\', \'" + databaseValues.visitee + "\', \'" + databaseValues.chaperone + "\');", connection))
                    {
                        connection.Open();
                        try
                        {
                            //command.Parameters.AddWithValue("@timeStamp", databaseValues.timeStamp.ToString());
                            //command.Parameters.AddWithValue("@nameString", databaseValues.name);
                            //command.Parameters.AddWithValue("@companyString", databaseValues.companyName);
                            //command.Parameters.AddWithValue("@citizen", databaseValues.citizenOrResident);
                            //command.Parameters.AddWithValue("@visitee", databaseValues.visitee);
                            //command.Parameters.AddWithValue("@otherPerson", databaseValues.otherPerson);
                            //command.Parameters.AddWithValue("@chaperone", databaseValues.chaperone);

                            command.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(string.Format("Database transfer error. Contact IT support.\n\n{0}", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error));
                            Application.Exit();
                            Environment.Exit(1);
                        }
                        //
                        // UPDATE END DATE
                        //
                        endDateTimePicker.ValueChanged -= this.endDateTimePicker_ValueChanged;
                        endDateTimePicker.MaxDate = databaseValues.timeStamp;
                        endDateTimePicker.Value = databaseValues.timeStamp;
                        endDateTimePicker.ValueChanged += this.endDateTimePicker_ValueChanged;

                        connection.Close();
                    }


                    // Go back to the intro screen
                    sigPlusNET1.SetTabletState(1);
                    sigPlusNET1.SetJustifyMode(oldJustifyMode);
                    sigPlusNET1.ClearSigWindow(1);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128); //Refresh entire tablet
                                                               //
                                                               // Load foreground image
                                                               //
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "            Welcome to ATI\n       Tap anywhere to begin");
                    sigPlusNET1.LCDSendGraphic(0, 2, 58, 45, atiLogo);
                    sigPlusNET1.KeyPadClearHotSpotList();
                    //
                    // SET UP HOT SPOTS
                    // Using LCD Coordinates
                    //
                    sigPlusNET1.KeyPadAddHotSpot(0, 1, 0, 0, 240, 128); // Hot spot for first screen "tap anywhere on screen to start"

                    sigPlusNET1.SetLCDCaptureMode(2);

                    ReloadDatabase();

                    // TO HERE
                }
                else
                {
                    Font please = new System.Drawing.Font("Arial", 16.0F, System.Drawing.FontStyle.Regular);
                    sigPlusNET1.LCDWriteString(0, 2, 55, 38, please, "Please Sign");
                    sigPlusNET1.LCDWriteString(0, 2, 30, 63, please, "Before Continuing...");
                    System.Threading.Thread.Sleep(2500);
                    sigPlusNET1.ClearTablet();
                    sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
                    sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "Enter Full Name:");
                    sigPlusNET1.LCDSendGraphic(0, 2, 0, 51, sigField);
                    sigPlusNET1.SetLCDCaptureMode(2);
                }
            }
            sigPlusNET1.ClearSigWindow(1);
        }

        private void startDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (startDateTimePicker.Value > endDateTimePicker.Value)
            {
                startDateTimePicker.ValueChanged -= this.startDateTimePicker_ValueChanged;
                startDateTimePicker.Value = endDateTimePicker.Value;
                startDateTimePicker.ValueChanged += this.startDateTimePicker_ValueChanged;
            }
            ReloadDatabase();
        }

        private void endDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (startDateTimePicker.Value > endDateTimePicker.Value)
            {
                endDateTimePicker.ValueChanged -= this.endDateTimePicker_ValueChanged;
                endDateTimePicker.Value = startDateTimePicker.Value;
                endDateTimePicker.ValueChanged += this.endDateTimePicker_ValueChanged;
            }
            ReloadDatabase();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //reset lcd to default settings
            sigPlusNET1.SetTabletState(1);
            sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128); //Clears entire LCD screen
            sigPlusNET1.LCDSetWindow(0, 0, 240, 128);
            sigPlusNET1.SetSigWindow(1, 0, 0, 240, 128);
            sigPlusNET1.KeyPadClearHotSpotList();
            sigPlusNET1.SetLCDCaptureMode(1); //Resets regular auto-clear inking
            sigPlusNET1.SetTabletState(0);
        }

        private void repetitiveMethod()
        {

            sigPlusNET1.SetLCDCaptureMode(2);

            // sleep for a second so user can see that checkbox was checked off
            Thread.Sleep(1000);

            // Clear window
            sigPlusNET1.ClearSigWindow(1);
            sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128);
            Font thankyou = new System.Drawing.Font("Arial", 16.0F, System.Drawing.FontStyle.Regular);
            sigPlusNET1.LCDWriteString(0, 2, 4, 10, thankyou, "Thank You For Signing");
            sigPlusNET1.LCDSendGraphic(0, 2, 58, 45, atiLogo);
            sigPlusNET1.SetTabletState(0); //turn off tablet to use justification below
            int oldJustifyMode = sigPlusNET1.GetJustifyMode();
            sigPlusNET1.SetJustifyMode(5); //this will zoom signature & justify to center
            Thread.Sleep(500);

            databaseValues.chaperone = string.Empty;

            // if not a citizen, request a chaperone
            if (!databaseValues.citizenOrResident)
                databaseValues.chaperone = Microsoft.VisualBasic.Interaction.InputBox("Visitor is not a citizen. Please assign a chaperone.", "Chaperone Text", "Enter chaperone's name here...");

            databaseValues.timeStamp = DateTime.Now;

            // CUT FROM HERE
            using (OleDbConnection connection = new OleDbConnection("Provider=sqloledb;Data Source=" + this.server + ";Initial Catalog=" + this.database + ";User Id=" + this.username + ";Password=" + this.password + ";"))
            using (OleDbCommand command = new OleDbCommand("INSERT INTO ATI_SignIn.dbo.VisitorsTable ([TimeStamp], [NameByteString], [CompanyRepresentedByteStream], [CitizenOrResident], [Visitee], [Chaperone]) VALUES (\'"+ databaseValues.timeStamp.ToString() + "\', \'"+ databaseValues.name + "\', \'" + databaseValues.companyName + "\', \'"+ databaseValues.citizenOrResident + "\', \'"+databaseValues.visitee+"\', \'"+ databaseValues.chaperone + "\');", connection))
            {
                connection.Open();
                try
                {
                    //command.Parameters.AddWithValue("@timeStamp", databaseValues.timeStamp.ToString());
                    //command.Parameters.AddWithValue("@nameString", databaseValues.name);
                    //command.Parameters.AddWithValue("@companyString", databaseValues.companyName);
                    //command.Parameters.AddWithValue("@citizen", databaseValues.citizenOrResident);
                    //command.Parameters.AddWithValue("@visitee", databaseValues.visitee);
                    //command.Parameters.AddWithValue("@chaperone", databaseValues.chaperone);

                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Database transfer error. Contact IT support.\n\n{0}", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error));
                    Application.Exit();
                    Environment.Exit(1);
                }
                //
                // UPDATE END DATE
                //
                endDateTimePicker.ValueChanged -= this.endDateTimePicker_ValueChanged;
                endDateTimePicker.MaxDate = databaseValues.timeStamp;
                endDateTimePicker.Value = databaseValues.timeStamp;
                endDateTimePicker.ValueChanged += this.endDateTimePicker_ValueChanged;

                connection.Close();
            }


            // Go back to the intro screen
            sigPlusNET1.SetTabletState(1);
            sigPlusNET1.SetJustifyMode(oldJustifyMode);
            sigPlusNET1.ClearSigWindow(1);
            sigPlusNET1.ClearTablet();
            sigPlusNET1.LCDRefresh(0, 0, 0, 240, 128); //Refresh entire tablet
                                                       //
                                                       // Load foreground image
                                                       //
            sigPlusNET1.LCDWriteString(0, 2, 4, 10, textFont, "            Welcome to ATI\n       Tap anywhere to begin");
            sigPlusNET1.LCDSendGraphic(0, 2, 58, 45, atiLogo);
            sigPlusNET1.KeyPadClearHotSpotList();
            //
            // SET UP HOT SPOTS
            // Using LCD Coordinates
            //
            sigPlusNET1.KeyPadAddHotSpot(0, 1, 0, 0, 240, 128); // Hot spot for first screen "tap anywhere on screen to start"

            sigPlusNET1.SetLCDCaptureMode(2);

            ReloadDatabase();

            // TO HERE
        }
    }
}