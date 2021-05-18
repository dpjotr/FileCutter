using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileCutter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            progressBar1.Visible = false;
            progressBar1.Maximum = 15;

            List<string> lines = new List<string>();

            System.IO.StreamReader file = 
                new System.IO.StreamReader(Application.StartupPath+"\\conf.txt");           
            for (int i=0; i<3; i++)
            {
                    lines.Add(file.ReadLine());
            }
            file.Close();

            string source = lines[0];
            string reference = lines[1];
            string destination = lines[2];

            textBox1.Text = source;
            textBox2.Text = reference;
            textBox3.Text = destination;

        }

        private async void button4_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Value = 1;
            
            ExcelCooperator ecoop = new ExcelCooperator(textBox1.Text, textBox2.Text, textBox3.Text);
            
            var start = DateTime.Now;
         
            Task<List<ClosedXML.Excel.IXLCell>> importedPartsCellsTask;
            try 
	        {
                importedPartsCellsTask = Task.Run(()=>ecoop.getCells(ecoop.getRows(ecoop.xlRefFilename, 18), 2));
	        }
	        catch (Exception exc)
	        {
                MessageBox.Show($"Unable to open reference file, set proper path " + exc.Message);
                return;
	        }

            
            Task<List<ClosedXML.Excel.IXLRow>> allSparePartsRowsTask = null;
            try
            {
                allSparePartsRowsTask = Task.Run(()=>ecoop.getRows(ecoop.xlSourceFileName, 2));
            }
            catch (Exception exc)
            {
                MessageBox.Show($"Unable to locate source file, set proper path " + exc.Message);
                return;
            }

            List<ClosedXML.Excel.IXLCell> importedPartsCells = await importedPartsCellsTask;
            List<ClosedXML.Excel.IXLRow> allSparePartsRows = await allSparePartsRowsTask;

            progressBar1.Value = 7;

            var finish = DateTime.Now;
            var taskDuration = (finish - start).TotalSeconds;

            richTextBox1
                .AppendText($@"Files were read, task took {taskDuration} seconds, {importedPartsCells.Last()}" + "\n");

            start = DateTime.Now;
            allSparePartsRows = allSparePartsRows
                .AsParallel()
                .Where(x => x.Cell(5).Value.ToString() != "311" && x.Cell(5).Value.ToString() != "312").ToList()
                ;
            finish = DateTime.Now;
            taskDuration = (finish - start).TotalSeconds;
            richTextBox1
                .AppendText($@"Spare parts rows list created, task took {taskDuration} seconds, {allSparePartsRows.Last()}" + "\n");
            progressBar1.Value = 8;

            start = DateTime.Now;
            var preResultList = allSparePartsRows
                .AsParallel()
                .Where(x => importedPartsCells.Select(y => y.Value).Contains(x.Cell(1).Value))
                .ToList();
            
            allSparePartsRows = null;
            finish = DateTime.Now;
            taskDuration = (finish - start).TotalSeconds;
            richTextBox1
                .AppendText($@"Result list created, task took {taskDuration} seconds, {preResultList.Last()}" + "\n");
            progressBar1.Value = 10;
            start = DateTime.Now;

            int fileNumber = 0;

            List<Task> files = new List<Task>();

            foreach (var x in ecoop.divideList(preResultList, (int)numericUpDown1.Value))
            {
                try
                {
                    files.Add(Task.Run(()=>ecoop.createFile(x, ++fileNumber)));
                }
                catch (Exception exc)
                {
                    MessageBox.Show($"Unable to create destination folder, set proper path" + exc.Message);
                    return;
                }
              
                richTextBox1.AppendText($"File {fileNumber} is being created\n");
            }

            foreach (var createFileTask in files) await createFileTask;
          
            preResultList = null;
            finish = DateTime.Now;
            taskDuration = (finish - start).TotalSeconds;
            richTextBox1
                .AppendText($"File(s) has been created, task took {taskDuration} seconds\n");
            
            File.WriteAllText(Application.StartupPath + "\\conf.txt", "");

            using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(Application.StartupPath + "\\conf.txt"))
            {
                file.WriteLine(textBox1.Text);
                file.WriteLine(textBox2.Text);
                file.WriteLine(textBox3.Text);
            }

            progressBar1.Value = 15;

            //file.Close();

            ecoop = null;
            progressBar1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    textBox1.Text = filePath;
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    textBox2.Text = filePath;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var folderPath = string.Empty;

            using (FolderBrowserDialog openFileFolder = new FolderBrowserDialog())
            {
                openFileFolder.Description = "Open folder";
                openFileFolder.ShowNewFolderButton = false;
                openFileFolder.RootFolder = Environment.SpecialFolder.Desktop;

                DialogResult result = openFileFolder.ShowDialog();
                if (result == DialogResult.OK)
                {
                    folderPath = openFileFolder.SelectedPath;
                    textBox3.Text = folderPath;
                }                
            }
            
        }


    }
}
