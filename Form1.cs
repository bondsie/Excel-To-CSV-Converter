using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel_To_CSV_Converter
{
    public partial class Form1 : Form
    {
        private string newName;

        public Form1()
        {
            InitializeComponent();
        }

        private void browse_Click(object sender, EventArgs e)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // defines the parameters of the file browsing fucntion. 
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.Title = "Browse Files";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel Files (.xlsx)|*.xlsx"; // searches specifically for applicable file types
            openFileDialog1.InitialDirectory = path + "\\Downloads\\"; //Sets default to the local download folder
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;

            if ((this.textBox1.Text != null) && (this.textBox1.Text != ""))
            {
                if (MessageBox.Show("Would you like to set saved path name to default?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                    var extension = Path.GetExtension(this.textBox1.Text);
                    if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) || string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                        this.textBox2.Text = this.textBox1.Text.Replace(extension, ".csv");
                    } else
                    {
                        MessageBox.Show("Invalid Path", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void convert_Click(object sender, EventArgs e)
        {
            if ((this.textBox1.Text != null) && (File.Exists(this.textBox1.Text)))
            {
                // this section grabs runs a test on the textbox for save as, sets a default if applicable (left blank)
                if (this.textBox2.Text == null)
                {
                    DateTime today = DateTime.Today;
                    string date = today.ToString("MM/dd/yyyy");
                    date = date.Replace("/", "");
                    string final = "LifeTest" + date + ".50093.csv";
                    newName = this.textBox1.Text.Replace("MetLife.xlsx", final);
                } else
                {
                    newName = this.textBox2.Text;
                }
                Console.WriteLine(newName);

                // get the value in the save as text box to write the file.
                if (this.textBox2.Text == string.Empty || Directory.Exists(Path.GetDirectoryName(this.textBox2.Text)))
                {
                    try
                    {
                        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(this.textBox1.Text);
                        workbook.SaveAs(newName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);
                        workbook.Close();
                        app.Quit();
                        Process.Start(Path.GetDirectoryName(this.textBox2.Text));

                    } catch
                    {
                        MessageBox.Show("No action completed.");
                    }
                    
                } else
                {
                    MessageBox.Show("Save file directory does not exist!");
                }
            } else
            {
                MessageBox.Show("File not found.");
            }
        }

        // clears out all value in both text boxes - button. 
        private void clear_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = string.Empty;
            this.textBox2.Text = string.Empty;
        }

        #region helpers
        public string fileName(object sender, EventArgs e)
        {
            string newName = this.textBox1.Text.Replace(".xlsx", ".csv");
            return newName;
        }

        #endregion
    }
}
