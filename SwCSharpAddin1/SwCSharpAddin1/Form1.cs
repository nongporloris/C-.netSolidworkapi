using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;

namespace SwCSharpAddin1
{
    public partial class Form1 : Form
    {
        string gb = "";

        ModelDoc2 doc;
        int fileerror;
        int filewarning;

        /***************/
        Feature swFeature;
        SelectionMgr swSelectionManager;
        Dimension swDim;
        string fileName;
        bool boolstatus;
        int errors;
        int warnings;



        public Form1(string input)
        {
            InitializeComponent();
        }
        public string MyProperty { get; set; }
        public ISldWorks swapp { get; set; }
        public int checkClick { get; private set; }

        public string name { get; set; }

        public static string myVal = "";

        public bool button1WasClicked = false;
        public string materialName {get; set; }

        public string dimensionName { get; set; }
        
        public string[] diresult { get; set; }
        public string[] result { get; set; }
        public string[] diresultTemp2 { get; set; }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        

        private void Form1_Load(object sender, EventArgs e)
        {

            string StandardBoxTemp = null;
            checkClick = 0;

            /********************************************************************************************/
            // Get files from directory into standard box.
            /*******************************************************************************************/

            DirectoryInfo filepath = new DirectoryInfo(@"E:\internship\InternApplicad\3dpart magic");
            DirectoryInfo[] folders = filepath.GetDirectories();
            StandardBox.DataSource = folders;
            StandardBoxTemp = StandardBox.SelectedItem.ToString();

        }

        private void StandardBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            /******************************************************************************************/
            // Clear item from the other box if this standaara box is changed.
            /*********************************************************************************************/
            

            TypeBox.SelectedItem = null;
            SizeBox.SelectedItem = null;
            CodeBox.SelectedItem = null;
            TypeBox.Items.Clear();
            SizeBox.Items.Clear();
            CodeBox.Items.Clear();

            /*******************************************************************************************/
            // Condition checking of the item in Standardbox.
            /********************************************************************************************/

            if (string.Compare(StandardBox.SelectedItem.ToString(), "DIN") == 0)
            {

                DirectoryInfo filePathForTypeBox = new DirectoryInfo(@"E:\internship\InternApplicad\3dpart magic\DIN\Database");
                DirectoryInfo[] foldersForTypeBox = filePathForTypeBox.GetDirectories();
                TypeBox.DataSource = foldersForTypeBox;

                checkClick = 1;
            }

            

            
        }

        private void TypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SizeBox.Items.Clear();
            SizeBox.Text = "--Select--";
            CodeBox.Text = "--Select--";
            CodeBox.Items.Clear();
            
            /************************************************************************************************************/
            // Get files from the directory for type box.
            /**************************************************************************************************************/
            
            DirectoryInfo filePathForSizeBox = new DirectoryInfo(@"E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem);

            string[] xlsxFiles = System.IO.Directory.GetFiles(filePathForSizeBox.ToString(), "*.xlsx");
            foreach (string file in xlsxFiles)
            {
                    
                // Remove the directory from the string
                string filename = file.Substring(file.LastIndexOf(@"\") + 1);
                // Remove the extension from the filename
                name = filename.Substring(0, filename.LastIndexOf(@"."));

                // Add the name to the combo box

                SizeBox.Items.Add(name);
                   
            }

         

        }

        private void SizeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            CodeBox.Items.Clear();
            CodeBox.SelectedItem = null;

            

            try
            {
                //**************************************************************************************************
                // Connect to the Data sheet from the selected excel file.
                /**************************************************************************************************/
                OleDbConnection conn2 = new OleDbConnection();
                conn2.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
                conn2.Open();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Data$]", conn2);
                DataTable ds = new DataTable();
                da.Fill(ds);
                CodeBox.Text = "--Select--";

                var dict = new Dictionary<Guid, string>();

                for (int i = 0; i < ds.Rows.Count; i++)
                {
                    CodeBox.Items.Add(ds.Rows[i]["Code"] + "-"+ ds.Rows[i]["Name"]);
                }


                conn2.Close();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetBaseException().ToString(), "Error In Connection");
            }

            /************************************************************************************************************/
            // Show the image from the selected database file.
            /*********************************************************************************************************/
            string[] standardPath = Directory.GetFiles(@"E:\internship\InternApplicad\3dpart magic\Part_DIN_ISO_and_Database\Database\_Preview\Standard\", "*.bmp", SearchOption.AllDirectories);
            foreach (string standardTemp  in standardPath)
            {
                string a = Path.GetFileNameWithoutExtension(standardTemp);
                if (a == SizeBox.SelectedItem.ToString())
                {
                    pictureBox1.Image = new Bitmap(@"E:\internship\InternApplicad\3dpart magic\Part_DIN_ISO_and_Database\Database\_Preview\Standard\" + SizeBox.SelectedItem + ".bmp");
                }
            }


            string[] picturePath = Directory.GetFiles(@"E:\internship\InternApplicad\3dpart magic\Part_DIN_ISO_and_Database\Database\_Preview", "*.bmp", SearchOption.AllDirectories);
            foreach (string pictureTemp in picturePath)
            {
                string a = Path.GetFileNameWithoutExtension(pictureTemp);
                if (a == SizeBox.SelectedItem.ToString())
                {
                    pictureBox2.Image = new Bitmap(@"E:\internship\InternApplicad\3dpart magic\Part_DIN_ISO_and_Database\Database\_Preview\" + SizeBox.SelectedItem + ".bmp");
                }
            }

            /************************************************************************************************************/
            // Connect to the data sheet from the selected file.
            /************************************************************************************************************/
            OleDbConnection conn5 = new OleDbConnection();
            conn5.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
            conn5.Open();
            OleDbDataAdapter dacon = new OleDbDataAdapter("SELECT * FROM [Data$]", conn5);
            DataTable di = new DataTable();
            dacon.Fill(di);

            dataGridView1.DataSource = di;

        }


        private void button1_Click(object sender, EventArgs e)
        {
            
            /***************************************************************************************************************/
            // Open .sldprt file from the slected database file in solidworks application.
            /**************************************************************************************************************/
             swapp.Visible = true;

            
             Debug.Print("Current working directory is " + swapp.GetCurrentWorkingDirectory());

             doc = swapp.OpenDoc6(@"E:\internship\InternApplicad\3dpart magic\Part_DIN_ISO_and_Database\Template\"+ SizeBox.SelectedItem +".sldprt", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_AutoMissingConfig, "", ref fileerror, ref filewarning);

            
             Debug.Print("Current working directory is still " + swapp.GetCurrentWorkingDirectory());

            
             swapp.SetCurrentWorkingDirectory(doc.GetPathName().Substring(0, doc.GetPathName().LastIndexOf(@"\")));

             Debug.Print("Current working directory is now " + swapp.GetCurrentWorkingDirectory());

            /*********************************************************************************************************************/

            string[] codeBoxItem = CodeBox.SelectedItem.ToString().Split('-');



            /***************************************************************************************************************/
            // Open .sldprt file from the slected database file in solidworks application.
            /**************************************************************************************************************/

            OleDbConnection conn3 = new OleDbConnection();
            conn3.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
            conn3.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Data$]", conn3);
            DataTable ds = new DataTable();
            da.Fill(ds);

            for (int i = 0; i < ds.Rows.Count; i++)
            {

                if (ds.Rows[i]["Code"].ToString() == (codeBoxItem[0]) & ds.Rows[i]["Name"].ToString() == (codeBoxItem[1]))
                {
                    /*******************************************************************************************************************/
                    // Connect the Dimension sheet from the selected file.
                    /*******************************************************************************************************************/
                    OleDbConnection conn = new OleDbConnection();

                    conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
                    conn.Open();
                    OleDbDataAdapter dimension = new OleDbDataAdapter("SELECT * FROM [Dimension$]", conn);
                    DataTable dimensionTable = new DataTable();
                    dimension.Fill(dimensionTable);


                    /********************************************************************************************************************/
                    // Connect the Data sheet from the slected file.
                    /*******************************************************************************************************************/

                    OleDbConnection conn4 = new OleDbConnection();
                    conn4.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
                    conn4.Open();
                    OleDbDataAdapter dacon = new OleDbDataAdapter("SELECT * FROM [Data$]", conn4);
                    DataTable di = new DataTable();
                    dacon.Fill(di);

                 /********************************************************************************************************************/
                 // get the data from the correct cell.
                    /******************************************************************************************************************/

                    foreach (DataColumn column in dimensionTable.Columns)
                    {
                        int o = 0;
                        result = new string[dimensionTable.Columns.Count];
                        result[o] = column.ColumnName;

                        foreach (DataColumn Datacolumn in di.Columns)
                        {
                            
                            diresult = new string[di.Columns.Count];
                            diresult[o] = Datacolumn.ColumnName;


                        

                            if (diresult[o] == result[o])
                            {
                                

                                for (int k = 0; k < dimensionTable.Rows.Count; k++)
                                {
                                    
                                    dimensionName = dimensionTable.Rows[k][diresult[o]].ToString();

                                    string tempValue = di.Rows[i][diresult[o]].ToString();
                                    
                                    /***************************************************************************************************/
                                    // Put the data to the selected dimensions in mm. unit.
                                    /****************************************************************************************************/
                                    if (dimensionName.Contains("Sketch") == true || dimensionName.Contains("@") == true || dimensionName.Contains("D") == true)
                                    {

                                        Dimension myDimension = default(Dimension);
                                        myDimension = (Dimension)doc.Parameter(dimensionName);
                                        myDimension.SystemValue = Convert.ToDouble(tempValue) * 0.001;

                                    }
                                    
                                    /******************************************************************************************************************/
                                    // Check to supress or unsupress to each revolve.
                                    /******************************************************************************************************************/
                                    if (dimensionName.Contains("Revolve") == true)
                                    {
                                        PartDoc part = (PartDoc)doc;
                                        IFeature swfeat = part.FeatureByName(dimensionName);
                                        swfeat.Select(false);

                                        if (tempValue == "1")
                                        {
                                            doc.EditUnsuppress();
                                        }
                                        else
                                        {
                                            doc.EditSuppress();
                                        }
                                    }
                                }



                            }


                        }


                        o++;
                    }

                    conn.Close();
                    conn4.Close();
                }

            }

            conn3.Close();

            /**********************************************************************************************************************/
            // open connection for the setting sheet from the selected file.
            /**********************************************************************************************************************/
            OleDbConnection settingconn = new OleDbConnection();
            settingconn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\internship\InternApplicad\3dpart magic\DIN\Database\" + TypeBox.SelectedItem + @"\" + SizeBox.SelectedItem + ".xlsx" + "; Extended Properties=Excel 12.0 XML;";
            settingconn.Open();
            OleDbDataAdapter settingcon = new OleDbDataAdapter("SELECT * FROM [Setting$]", settingconn);
            DataTable set = new DataTable();
            settingcon.Fill(set);

            /********************************************************************************************************************/
            // Get the material form the setting sheet.
            /*******************************************************************************************************************/
            for (int setting = 0; setting < set.Rows.Count; setting++)
            {
                materialName = set.Rows[setting]["MaterialType"].ToString();
            }

            settingconn.Close();

         /**************************************************************************************************************************/
         // Set the material to the opened file follow by the material from setting sheet.
            /***********************************************************************************************************************/

            object[] vMatDBarr = null;
            object[] vMatDB = null;
            PartDoc swPart = (PartDoc)doc;
            Body2 swBody = default(Body2);
            int ipre = 0;
            int jpre = 0;
            object[] Bodies = null;
            long BodyMaterialError = 0;
            string sMatName = "";
            string sMatDB = "";
            bool boolstat = false;

            vMatDBarr = (object[])swapp.GetMaterialDatabases();
            Debug.Print("Material schema pathname = " + swapp.GetMaterialSchemaPathName());

            for (ipre = 0; ipre < vMatDBarr.Length; ipre++)
            {
                Debug.Print(" Material database: " + vMatDB);
            }
            Debug.Print("");

            Bodies = (object[])swPart.GetBodies2((int)swBodyType_e.swAllBodies, false);

            for (jpre = 0; jpre < Bodies.Length; jpre++)
            {
                swBody = (Body2)Bodies[jpre];

                Debug.Print(swBody.Name);

                swBody.Select2(false, null);
                BodyMaterialError = swBody.SetMaterialProperty("Default", @"C:/Program Files/SOLIDWORKS Corp/SOLIDWORKS/lang/english/sldmaterials/SOLIDWORKS Materials.sldmat", materialName);
                sMatName = swBody.GetMaterialPropertyName("", out sMatDB);

                if (string.IsNullOrEmpty(sMatName))
                {
                    Debug.Print("Body " + jpre + "'s material name: No material applied");
                }
                else
                {
                    Debug.Print("Body " + jpre + "'s material name: " + sMatName);
                    Debug.Print("Body " + jpre + "'s material database: " + sMatDB);
                    Debug.Print(" ");
                }
            }

            /*******************************************************************************************************************/
            // Rebuild and save as file.
            /*********************************************************************************************************************/

            doc.EditRebuild3();
           
            swapp.RunCommand((int)swCommands_e.swCommands_SaveAs, "");
           
            /*******************************************************************************************************************/            
        }

        private void CodeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
          

        }
    }
}
