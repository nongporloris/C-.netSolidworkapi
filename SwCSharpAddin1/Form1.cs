using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Drawing;
using GemBox.Spreadsheet;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace app1
{
    public partial class Form1 : Form
    {
        private OleDbCommand sql;

        public int checkValue { get; private set; }

        public Form1()
        {
            InitializeComponent();
            
        }

        public interface Workbook : IDisposable, System.ComponentModel.IComponent, System.ComponentModel.ISupportInitialize, System.Windows.Forms.IBindableComponent
        { 

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\internship\InternApplicad\Database\Bearings\DIN 720-1.mdb");
            con.Open();

            OleDbCommand cmd = new OleDbCommand("select * from Data", con);
            OleDbDataReader rd = cmd.ExecuteReader();
            DataTable dt = new DataTable();

            dt.Load(rd);


            OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
            OleDbDataReader rd1 = cmd1.ExecuteReader();
            DataTable dt1 = new DataTable();

            dt1.Load(rd1);


            OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
            OleDbDataReader rd2 = cmd2.ExecuteReader();
            DataTable dt2 = new DataTable();

            dt2.Load(rd2);
            con.Close();
            dataGridView1.DataSource = dt;
            dataGridView2.DataSource = dt1;
            dataGridView3.DataSource = dt2;
        }

      

        private void button2_Click_1(object sender, EventArgs e)
        {

            /*          OpenFileDialog openFileDialog1 = new OpenFileDialog
                        {
                            InitialDirectory = @"E:\internship\InternApplicad\Database",
                            Title = "Browse mdb Files",

                            CheckFileExists = true,
                            CheckPathExists = true,

                            DefaultExt = "txt",
                            Filter = "folder files (*.bat)|*.bat",
                            FilterIndex = 2,
                            RestoreDirectory = true,

                            ReadOnlyChecked = true,
                            ShowReadOnly = true
                        };

                        if (openFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            textBox1.Text = openFileDialog1.FileName;

                            string text1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                            string text2 = text1 + textBox1.Text;

                            OleDbConnection con = new OleDbConnection(@text2);
                            con.Open();

                            OleDbCommand cmd = new OleDbCommand("select * from Data", con);
                            OleDbDataReader rd = cmd.ExecuteReader();
                            DataTable dt = new DataTable();

                            dt.Load(rd);


                            OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
                            OleDbDataReader rd1 = cmd1.ExecuteReader();
                            DataTable dt1 = new DataTable();

                            dt1.Load(rd1);


                            OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
                            OleDbDataReader rd2 = cmd2.ExecuteReader();
                            DataTable dt2 = new DataTable();

                            dt2.Load(rd2);
                            con.Close();
                            dataGridView1.DataSource = dt;
                            dataGridView2.DataSource = dt1;
                            dataGridView3.DataSource = dt2;
                        }
            */

            if (checkValue != 1)
            {
                MessageBox.Show("Please input the direction path");
                return;
            }
            FolderBrowserDialog FBD = new FolderBrowserDialog();

            if (FBD.ShowDialog()==DialogResult.OK)
            {
                listBox1.Items.Clear();
                string[] files = Directory.GetFiles(FBD.SelectedPath);
                string[] dirs = Directory.GetDirectories(FBD.SelectedPath);

                foreach (string file in files)
                {
                    listBox1.Items.Add(file);
                }
                foreach (string dir in dirs)
                {
                    listBox1.Items.Add(dir);
                }

            }



            var list = new List<string>();

            foreach (var item in listBox1.Items)
            {
                list.Add(item.ToString());
                
            }

            for (int a = 0; a < listBox1.Items.Count; a++)
            {

                textBox1.Text = list[a];

                string text1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                string text2 = text1 + textBox1.Text;

                OleDbConnection con = new OleDbConnection(@text2);
                con.Open();

                OleDbCommand cmd = new OleDbCommand("select * from Data", con);
                OleDbDataReader rd = cmd.ExecuteReader();
                DataTable dt = new DataTable();

                dt.Load(rd);


                OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
                OleDbDataReader rd1 = cmd1.ExecuteReader();
                DataTable dt1 = new DataTable();

                dt1.Load(rd1);


                OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
                OleDbDataReader rd2 = cmd2.ExecuteReader();
                DataTable dt2 = new DataTable();

                dt2.Load(rd2);
                con.Close();
                dataGridView1.DataSource = dt;
                dataGridView2.DataSource = dt1;
                dataGridView3.DataSource = dt2;

                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                Microsoft.Office.Interop.Excel._Worksheet worksheet2 = null;
                Microsoft.Office.Interop.Excel._Worksheet worksheet3 = null;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Data";

                int count = workbook.Worksheets.Count;
                Excel.Worksheet addedSheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);

                worksheet2 = workbook.Sheets["Sheet2"];
                worksheet2 = workbook.ActiveSheet;
                worksheet2.Name = "Dimension";

                int count1 = workbook.Worksheets.Count;
                Excel.Worksheet addedSheet1 = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);
                worksheet3 = workbook.Sheets["Sheet3"];
                worksheet3 = workbook.ActiveSheet;
                worksheet3.Name = "Setting";

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }





                for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    worksheet2.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        if (dataGridView2.Rows[i].Cells[j].Value != null)
                        {
                            worksheet2.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }






                for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
                {
                    worksheet3.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.Columns.Count; j++)
                    {
                        if (dataGridView3.Rows[i].Cells[j].Value != null)
                        {
                            worksheet3.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }



                string filenameoutput = Path.GetFileNameWithoutExtension(textBox1.Text);

                workbook.SaveAs(@textBox2.Text + @"\" + filenameoutput + ".xlsx");
              



 /*               var saveFileDialog = new SaveFileDialog();
                saveFileDialog.FileName = "output";
                saveFileDialog.DefaultExt = ".xlsx";
                

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                   workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
               

*/




                app.Quit();


            }

            MessageBox.Show("Convert Complete");


            /*           MessageBox.Show(list[0]);
                       MessageBox.Show(list[1]);
                       MessageBox.Show(list[2]);
                       MessageBox.Show(list[3]);
                       MessageBox.Show(list[4]);
            */
            /*           for (int a = 0; a < listBox1.Items.Count; a++)
                       {
                           listBox1.Items.Add(listBox1.Items[a].ToString());

                           MessageBox.Show((string)listBox1.Items[a]);
                       }

           */

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            /*           OleDbConnection con = new OleDbConnection();
                       string dbProvider = null;
                       string dbSource = null;
                       string dbTableName = null;
                       string sql1 = null;
                       int MaxRows = 0;
                       string outputCSV = "";

                       dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source = ";
                       dbSource = textBox1.Text;
                       dbTableName = "Data";
                       int noOfColumns = 1;

                       con.ConnectionString = dbProvider + dbSource;
                       DataSet ds = new DataSet();
                       OleDbDataAdapter da;
                       sql1 = "SELECT * FROM " + dbTableName;

                       con.Open();
                       da = new OleDbDataAdapter(sql, con);
                       da.Fill(ds, dbTableName);
                       MaxRows = ds.Tables("Data").Rows.Count;
                       con.Close();

                       for (int i = 0; i < MaxRows; i++)
                       {
                           for (int j = 0; j < noOfColumns; j++)
                           {


                               outputCSV = outputCSV + ds.Tables("FancyName").Rows(i).Item(j) + ",";

                           }
                           outputCSV = outputCSV + "\r\n";
                       }*/

 /*           int count_row = dataGridView1.RowCount;
            int count_cell = dataGridView1.Rows[0].Cells.Count;

            int count_row2 = dataGridView2.RowCount;
            int count_cell2 = dataGridView2.Rows[0].Cells.Count;

            int count_row3 = dataGridView3.RowCount;
            int count_cell3 = dataGridView3.Rows[0].Cells.Count;

            for (int row_index = 0; row_index <= count_row - 2; row_index++)
            {
                for (int cell_index = 0; cell_index <= count_cell - 1; cell_index++)
                {
                    //MessageBox.Show(dataGridView1.Rows[row_index].Cells[cell_index].Value.ToString());
                    text_box_export.Text = text_box_export.Text + dataGridView1.Rows[row_index].Cells[cell_index].Value.ToString() +",";
                    //text_box_export2.Text = text_box_export2.Text + dataGridView2.Rows[row_index].Cells[cell_index].Value.ToString() + ",";
                    //text_box_export3.Text = text_box_export3.Text + dataGridView3.Rows[row_index].Cells[cell_index].Value.ToString() + ",";
                }
                text_box_export.Text = text_box_export.Text + "\r\n";
                //text_box_export2.Text = text_box_export2.Text + "\r\n";
                //text_box_export3.Text = text_box_export3.Text + "\r\n";
            }
            for (int row_index2 = 0; row_index2 <= count_row2 - 2; row_index2++)
            {
                for (int cell_index2 = 0; cell_index2 <= count_cell2 - 1; cell_index2++)
                {
                    //MessageBox.Show(dataGridView1.Rows[row_index].Cells[cell_index].Value.ToString());
                   
                    text_box_export2.Text = text_box_export2.Text + dataGridView2.Rows[row_index2].Cells[cell_index2].Value.ToString() + ",";
                   
                }
                text_box_export2.Text = text_box_export2.Text + "\r\n";
            }
            for (int row_index3 = 0; row_index3 <= count_row3 - 2; row_index3++)
            {
                for (int cell_index3 = 0; cell_index3 <= count_cell3 - 1; cell_index3++)
                {
                    //MessageBox.Show(dataGridView1.Rows[row_index].Cells[cell_index].Value.ToString());

                    text_box_export3.Text = text_box_export3.Text + dataGridView3.Rows[row_index3].Cells[cell_index3].Value.ToString() + ",";

                }
                text_box_export3.Text = text_box_export3.Text + "\r\n";
            }

            exportSheet.Text = text_box_export.Text + text_box_export2.Text + text_box_export3.Text; 

            System.IO.File.WriteAllText(@"E:\internship\InternApplicad\app1\exportSheet1.csv", text_box_export.Text);
            System.IO.File.WriteAllText(@"E:\internship\InternApplicad\app1\exportSheet2.csv", text_box_export2.Text);
            System.IO.File.WriteAllText(@"E:\internship\InternApplicad\app1\exportSheet3.csv", text_box_export3.Text);

*/



            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            Microsoft.Office.Interop.Excel._Worksheet worksheet2 = null;
            Microsoft.Office.Interop.Excel._Worksheet worksheet3 = null;

            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Data";

            int count = workbook.Worksheets.Count;
            Excel.Worksheet addedSheet = workbook.Worksheets.Add(Type.Missing,workbook.Worksheets[count], Type.Missing, Type.Missing);

            worksheet2 = workbook.Sheets["Sheet2"];
            worksheet2 = workbook.ActiveSheet;
            worksheet2.Name = "Dimension";

            int count1 = workbook.Worksheets.Count;
            Excel.Worksheet addedSheet1 = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);
            worksheet3 = workbook.Sheets["Sheet3"];
            worksheet3 = workbook.ActiveSheet;
            worksheet3.Name = "Setting";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1,i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i<dataGridView1.Rows.Count; i++)
            {
                for (int j =0; j<dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }





            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                worksheet2.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                    {
                        worksheet2.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
 





            for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
            {
                worksheet3.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                {
                    if (dataGridView3.Rows[i].Cells[j].Value != null)
                    {
                        worksheet3.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }







            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "Output";
            saveFileDialog.DefaultExt = ".xlsx";

            if (saveFileDialog.ShowDialog()==DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }







            app.Quit();




            MessageBox.Show("Convert Complete");
        }


        private void Form1_Load(object sender, EventArgs e)
        {
        
        }

        private void button4_Click(object sender, EventArgs e)
        {
            checkValue = 1;

            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                listBox2.Items.Clear();
                string[] files = Directory.GetFiles(FBD.SelectedPath,"*.mdb",SearchOption.AllDirectories);
                string[] dirs = Directory.GetDirectories(FBD.SelectedPath);

                foreach (string file in files)
                {
                    listBox2.Items.Add(file);
                }
                foreach (string dir in dirs)
                {
                    listBox2.Items.Add(dir);
                }

            }



            var list2 = new List<string>();
            string v;
            foreach (var item in listBox2.Items)
            {
                list2.Add(item.ToString());

            }

            textBox2.Text = Path.GetDirectoryName(list2[0].ToString());
            MessageBox.Show(textBox2.Text);
            
               
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"E:\internship\InternApplicad\Database",
                Title = "Browse mdb Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "txt",
                Filter = "folder files (*.mdb)|*.mdb",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;

                string text1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                string text2 = text1 + textBox1.Text;

                OleDbConnection con = new OleDbConnection(@text2);
                con.Open();

                OleDbCommand cmd = new OleDbCommand("select * from Data", con);
                OleDbDataReader rd = cmd.ExecuteReader();
                DataTable dt = new DataTable();

                dt.Load(rd);


                OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
                OleDbDataReader rd1 = cmd1.ExecuteReader();
                DataTable dt1 = new DataTable();

                dt1.Load(rd1);


                OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
                OleDbDataReader rd2 = cmd2.ExecuteReader();
                DataTable dt2 = new DataTable();

                dt2.Load(rd2);
                con.Close();
                dataGridView1.DataSource = dt;
                dataGridView2.DataSource = dt1;
                dataGridView3.DataSource = dt2;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {


            if (checkValue != 1)
            {
                MessageBox.Show("Please input the direction path");
                return;
            }


            Directory.CreateDirectory(@textBox2.Text + @"\Database");

            FolderBrowserDialog FBDatabase = new FolderBrowserDialog();

            if (FBDatabase.ShowDialog() == DialogResult.OK)
            {
                listBox3.Items.Clear();
                string[] filesDatabase = Directory.GetFiles(FBDatabase.SelectedPath,"*.mdb",SearchOption.AllDirectories);             //Database file
                string[] dirsDatabase = Directory.GetDirectories(FBDatabase.SelectedPath);

                foreach (string fileDatabase in filesDatabase)
                {
                    listBox3.Items.Add(fileDatabase);
                }
                foreach (string dirDatabase in dirsDatabase)
                {
                    listBox3.Items.Add(dirDatabase);
                }

            }

            var listDatabase = new List<string>();

            foreach (var itemDatabase in listBox3.Items)
            {
                listDatabase.Add(itemDatabase.ToString());
                
            }

            for (int a = 0; a < listBox3.Items.Count; a++)
            {

                textBox1.Text = listDatabase[a];

                

                string text1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                string text2 = text1 + textBox1.Text;

                if (textBox1.Text.Contains("FilePath"))
                {
                    a = a+1;
                }

                string folderNameTemp = Path.GetDirectoryName(listDatabase[a]);
                string folderName = Path.GetFileNameWithoutExtension(folderNameTemp);
                
                textBox1.Text = listDatabase[a];
                text2 = text1 + textBox1.Text;

                if (listDatabase[a].Contains(folderName))
                {
                    Directory.CreateDirectory(@textBox2.Text + @"\Database\"+ folderName);
                }

                if (folderName.Contains("Database") )
                {
                    MessageBox.Show("Convert batch file complete.");
                    return;
                }

                OleDbConnection con = new OleDbConnection(@text2);
                con.Open();

                OleDbCommand cmd = new OleDbCommand("select * from Data", con);
                OleDbDataReader rd = cmd.ExecuteReader();
                DataTable dt = new DataTable();

                dt.Load(rd);


                OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
                OleDbDataReader rd1 = cmd1.ExecuteReader();
                DataTable dt1 = new DataTable();

                dt1.Load(rd1);


                OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
                OleDbDataReader rd2 = cmd2.ExecuteReader();
                DataTable dt2 = new DataTable();

                dt2.Load(rd2);
                con.Close();
                dataGridView1.DataSource = dt;
                dataGridView2.DataSource = dt1;
                dataGridView3.DataSource = dt2;






                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                Microsoft.Office.Interop.Excel._Worksheet worksheet2 = null;
                Microsoft.Office.Interop.Excel._Worksheet worksheet3 = null;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Data";

                int count = workbook.Worksheets.Count;
                Excel.Worksheet addedSheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);

                worksheet2 = workbook.Sheets["Sheet2"];
                worksheet2 = workbook.ActiveSheet;
                worksheet2.Name = "Dimension";

                int count1 = workbook.Worksheets.Count;
                Excel.Worksheet addedSheet1 = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);
                worksheet3 = workbook.Sheets["Sheet3"];
                worksheet3 = workbook.ActiveSheet;
                worksheet3.Name = "Setting";

                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }





                for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    worksheet2.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        if (dataGridView2.Rows[i].Cells[j].Value != null)
                        {
                            worksheet2.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }






                for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
                {
                    worksheet3.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.Columns.Count; j++)
                    {
                        if (dataGridView3.Rows[i].Cells[j].Value != null)
                        {
                            worksheet3.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                }


                string filenameoutput = Path.GetFileNameWithoutExtension(textBox1.Text);

                workbook.SaveAs(@textBox2.Text + @"\Database" + @"\" + folderName + @"\" + filenameoutput + ".xlsx");
                

                app.Quit();
            }

            MessageBox.Show("Convert all file complete");



















            /*

                        FolderBrowserDialog FBD = new FolderBrowserDialog();

                        if (FBD.ShowDialog() == DialogResult.OK)
                        {
                            listBox1.Items.Clear();
                            string[] files = Directory.GetFiles(FBD.SelectedPath);
                            string[] dirs = Directory.GetDirectories(FBD.SelectedPath);

                            foreach (string file in files)
                            {
                                listBox1.Items.Add(file);
                            }
                            foreach (string dir in dirs)
                            {
                                listBox1.Items.Add(dir);
                            }

                        }



                        var list = new List<string>();

                        foreach (var item in listBox1.Items)
                        {
                            list.Add(item.ToString());

                        }

                        for (int a = 0; a < listBox1.Items.Count; a++)
                        {

                            textBox1.Text = list[a];


                            string text1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
                            string text2 = text1 + textBox1.Text;

                            OleDbConnection con = new OleDbConnection(@text2);
                            con.Open();

                            OleDbCommand cmd = new OleDbCommand("select * from Data", con);
                            OleDbDataReader rd = cmd.ExecuteReader();
                            DataTable dt = new DataTable();

                            dt.Load(rd);


                            OleDbCommand cmd1 = new OleDbCommand("select * from Dimension", con);
                            OleDbDataReader rd1 = cmd1.ExecuteReader();
                            DataTable dt1 = new DataTable();

                            dt1.Load(rd1);


                            OleDbCommand cmd2 = new OleDbCommand("select * from Setting", con);
                            OleDbDataReader rd2 = cmd2.ExecuteReader();
                            DataTable dt2 = new DataTable();

                            dt2.Load(rd2);
                            con.Close();
                            dataGridView1.DataSource = dt;
                            dataGridView2.DataSource = dt1;
                            dataGridView3.DataSource = dt2;






                            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                            Microsoft.Office.Interop.Excel._Worksheet worksheet2 = null;
                            Microsoft.Office.Interop.Excel._Worksheet worksheet3 = null;

                            worksheet = workbook.Sheets["Sheet1"];
                            worksheet = workbook.ActiveSheet;
                            worksheet.Name = "Data";

                            int count = workbook.Worksheets.Count;
                            Excel.Worksheet addedSheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);

                            worksheet2 = workbook.Sheets["Sheet2"];
                            worksheet2 = workbook.ActiveSheet;
                            worksheet2.Name = "Dimension";

                            int count1 = workbook.Worksheets.Count;
                            Excel.Worksheet addedSheet1 = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[count], Type.Missing, Type.Missing);
                            worksheet3 = workbook.Sheets["Sheet3"];
                            worksheet3 = workbook.ActiveSheet;
                            worksheet3.Name = "Setting";

                            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                            {
                                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                            }
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                                {
                                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                                    {
                                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                                    }
                                }
                            }





                            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                            {
                                worksheet2.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                            }
                            for (int i = 0; i < dataGridView2.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                                {
                                    if (dataGridView2.Rows[i].Cells[j].Value != null)
                                    {
                                        worksheet2.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                                    }
                                }
                            }






                            for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
                            {
                                worksheet3.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                            }
                            for (int i = 0; i < dataGridView3.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                                {
                                    if (dataGridView3.Rows[i].Cells[j].Value != null)
                                    {
                                        worksheet3.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                                    }
                                }
                            }

                            string filenameoutput2 = Path.GetFileNameWithoutExtension(textBox1.Text);


                            string filenameoutput3 = Path.GetDirectoryName(textBox1.Text);
            //                Directory.CreateDirectory(filenameoutput2);

                            string filenameoutput = Path.GetFileNameWithoutExtension(textBox1.Text);

                            workbook.SaveAs(@textBox2.Text+ @"\Database" + @"\" + filenameoutput + ".xlsx");
                            MessageBox.Show(textBox2.Text);

                            app.Quit();
                        }

                        MessageBox.Show("Convert all file complete");

                       */

        }
    }
}
