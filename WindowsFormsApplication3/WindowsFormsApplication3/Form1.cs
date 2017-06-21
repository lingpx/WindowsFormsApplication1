using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.IO.Compression;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        string sourcefolderpath { get; set; }
        string outputfolderpath { get; set; }
        string zipfolderpath { get; set; }
        string wordfolder { get; set; }

        
        string[] sourcefilenames { get; set; }
        string[] targetfilenames { get; set; }
        
       

        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }

        

        public Form1()
        {
            InitializeComponent();
            sourcefolderpath = @"N:\Power Engineering\04 Engineering Documents\94- Power System Engineering\94.11 Contractors\94.11.3 Test & Commissioning";
            label3.Text = sourcefolderpath;

            if (sourcefolderpath != "")
            {
                // Storing list of all docx files in sourcefilenames array
                string[] sourcefilenames = System.IO.Directory.GetFiles(sourcefolderpath, "*.docx").Select(path => System.IO.Path.GetFileName(path)).ToArray();

                for (int i = 0; i < sourcefilenames.Length; i++)
                {
                    checkedListBox1.Items.Insert(i, sourcefilenames[i]);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Directory.CreateDirectory(outputfolderpath + "\\Word");

            foreach (string fname in checkedListBox1.CheckedItems)
            {

                
                string fileName = fname;
                //string sourcePath = @"C:\Users\Public\TestFolder";
                string targetPath = outputfolderpath + "\\Word";
                //   MessageBox.Show("Output path set to: " + targetPath);

                string sourcePath = sourcefolderpath;
                //    MessageBox.Show("Source path set to: "+sourcePath);
                // string targetPath = (string)outputfolderpath;
                // MessageBox.Show(outputfolderpath);

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, fileName);

                // To copy a folder's contents to a new location:
                // Create a new target folder, if necessary.
                if (!System.IO.Directory.Exists(targetPath))
                {
                    System.IO.Directory.CreateDirectory(targetPath);
                }

                if (System.IO.File.Exists(sourceFile))
                {
                    // To copy a file to another location and 
                    // overwrite the destination file if it already exists.
                    listBox1.Items.Add(fname);
                    System.IO.File.Copy(sourceFile, destFile, true);

                    // To copy all the files in one directory to another directory.
                    // Get the files in the source folder. (To recursively iterate through
                    // all subfolders under the current directory, see
                    // "How to: Iterate Through a Directory Tree.")
                    // Note: Check for target path was performed previously
                    //       in this code example.
                    /*if (System.IO.Directory.Exists(sourcePath))
                    {
                        string[] files = System.IO.Directory.GetFiles(sourcePath);

                        // Copy the files and overwrite destination files if they already exist.
                        foreach (string s in files)
                        {
                            // Use static Path methods to extract only the file name from the path.
                            fileName = System.IO.Path.GetFileName(s);
                            destFile = System.IO.Path.Combine(targetPath, fileName);
                            System.IO.File.Copy(s, destFile, true);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Source path does not exist!");
                    }
                    */
                    // Keep console window open in debug mode.
                    //Console.WriteLine("Press any key to exit.");
                    //Console.ReadKey();

                }
                else
                    MessageBox.Show(sourceFile + " does not exist");
                
            }

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void SourceFolderButton_Click(object sender, EventArgs e)
        {
            
            sourcefolderpath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dr = fbd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                sourcefolderpath = fbd.SelectedPath;
                label3.Text = sourcefolderpath;
            }

            if (sourcefolderpath!="")
            {
                    // Storing list of all docx files in sourcefilenames array
                        string[] sourcefilenames = System.IO.Directory.GetFiles(sourcefolderpath, "*.docx").Select(path => System.IO.Path.GetFileName(path)).ToArray();

                        for (int i = 0; i < sourcefilenames.Length; i++)
                        {
                            checkedListBox1.Items.Insert(i, sourcefilenames[i]);
                        }
            }
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        public void OutputFolderButton_Click(object sender, EventArgs e)
        {
            outputfolderpath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dr = fbd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                outputfolderpath = fbd.SelectedPath;
                OutputLabel.Text = outputfolderpath;
            }

            wordfolder = outputfolderpath + "\\Word";
            zipfolderpath = outputfolderpath + "\\Zip";
            ZIPLabel.Text = zipfolderpath;
        }

        public void ZipFolderButton_Click(object sender, EventArgs e)
        {
            zipfolderpath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dr = fbd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                zipfolderpath = fbd.SelectedPath;
                ZIPLabel.Text = zipfolderpath;
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void complieAndEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void toolStripContainer2_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        // Convert to PDF button click
        private void button2_Click(object sender, EventArgs e)
        {
            string PDFFolder = outputfolderpath + "\\PDFs";
            
            Directory.CreateDirectory(PDFFolder);
            // Storing list of all docx files in targetfilenames array WITHOUT extension
            targetfilenames = System.IO.Directory.GetFiles(wordfolder + "\\", "*.docx").Select(path => System.IO.Path.GetFileNameWithoutExtension(path)).ToArray();
           

            // Looping through every docx document in wordfolder and converting to PDF and saving in PDFFolder
            for (int i = 0; i < targetfilenames.Length; i++)
            {
                
                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                wordDocument = appWord.Documents.Open(wordfolder + "\\" + targetfilenames[i]+".docx");
                wordDocument.ExportAsFixedFormat(PDFFolder + "\\"+ targetfilenames[i]+".pdf", WdExportFormat.wdExportFormatPDF);
                appWord.Quit(WdSaveOptions.wdDoNotSaveChanges);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(appWord);

                
            }


        }

        

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void ButtonZIP_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(zipfolderpath);
            string zipfile = zipfolderpath + "\\zipfile.zip";
            ZipFile.CreateFromDirectory(outputfolderpath+"\\PDFs", zipfile);
        
        }

        private void OutputLabel_Click(object sender, EventArgs e)
        {

        }

        private void ZIPFolder_Click(object sender, EventArgs e)
        {
            zipfolderpath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dr = fbd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                zipfolderpath = fbd.SelectedPath;
                ZIPLabel.Text = zipfolderpath;
            }
        }

        private void ZIPLabel_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
            
        }

        private void UpdateWordProperties_Click(object sender, EventArgs e)
        {


            Word.Application oWord;
            Word._Document oDoc;
            object oMissing = Missing.Value;
            object oDocCustomProps;
            //object fileName=filename;
            object fileName;
            string strIndex, strValue;
            string filename;

            // Storing list of all docx files in targetfilenames array WITH extension
            targetfilenames = System.IO.Directory.GetFiles(wordfolder + "\\", "*.docx").Select(path => System.IO.Path.GetFileName(path)).ToArray();

            // Looping through every docx document in wordfolder 
            for (int i = 0; i < targetfilenames.Length; i++)
            {

                filename = wordfolder + "\\" + targetfilenames[i]; //getting filename 
                fileName = filename; // Passing full filename into the object fileName

                //Create an instance of Microsoft Word and make it visible.
                oWord = new Word.Application();
                oWord.Visible = true;

                //Open the word document and get the CustomDocumentProperties collection.
                oDoc = oWord.Documents.Open(ref fileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                oDocCustomProps = oDoc.CustomDocumentProperties;
                Type typeDocCustomProps = oDocCustomProps.GetType();

                //Get the MT_TSP_ProjectCode property 
                strIndex = "MT_TSP_ProjectCode";

                object oDocMT_TSP_ProjectCodeProp = typeDocCustomProps.InvokeMember("Item",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocCustomProps,
                                           new object[] { strIndex });
                Type typeDocMT_TSP_ProjectCodeProp = oDocMT_TSP_ProjectCodeProp.GetType();
                strValue = typeDocMT_TSP_ProjectCodeProp.InvokeMember("Value",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocMT_TSP_ProjectCodeProp,
                                           new object[] { }).ToString();
                // MessageBox.Show("The MT_TSP_ProjectCode is: " + strValue, "MT_TSP_ProjectCode");

                // Set the MT_TSP_ProjectCode strindex to the text set in Project Code

                typeDocMT_TSP_ProjectCodeProp.InvokeMember("Item",
                    BindingFlags.Default |
                    BindingFlags.SetProperty,
                    null, oDocCustomProps,
                    new object[] { strIndex, textBox1.Text });


                //Get the MT_TSP_ProjectName property 
                strIndex = "MT_TSP_ProjectName";

                object oDocMT_TSP_ProjectNameProp = typeDocCustomProps.InvokeMember("Item",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocCustomProps,
                                           new object[] { strIndex });
                Type typeDocMT_TSP_ProjectNameProp = oDocMT_TSP_ProjectNameProp.GetType();
                strValue = typeDocMT_TSP_ProjectCodeProp.InvokeMember("Value",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocMT_TSP_ProjectNameProp,
                                           new object[] { }).ToString();
                
                // Set the MT_TSP_ProjectName strindex to the text set in Project Name

                typeDocMT_TSP_ProjectCodeProp.InvokeMember("Item",
                    BindingFlags.Default |
                    BindingFlags.SetProperty,
                    null, oDocCustomProps,
                    new object[] { strIndex, textBox2.Text });

                //Get the MT_TSP_Projectlocation property 
                strIndex = "MT_TSP_Projectlocation";

                object oDocMT_TSP_ProjectlocationProp = typeDocCustomProps.InvokeMember("Item",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocCustomProps,
                                           new object[] { strIndex });
                Type typeDocMT_TSP_ProjectlocationProp = oDocMT_TSP_ProjectlocationProp.GetType();
                strValue = typeDocMT_TSP_ProjectlocationProp.InvokeMember("Value",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocMT_TSP_ProjectlocationProp,
                                           new object[] { }).ToString();

                // Set the MT_TSP_Projectlocation strindex to the text set in Project Location

                typeDocMT_TSP_ProjectlocationProp.InvokeMember("Item",
                    BindingFlags.Default |
                    BindingFlags.SetProperty,
                    null, oDocCustomProps,
                    new object[] { strIndex, textBox3.Text });

                //Get the MT_TSP_Date property 
                strIndex = "MT_TSP_Date";

                object oDocMT_TSP_DateProp = typeDocCustomProps.InvokeMember("Item",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocCustomProps,
                                           new object[] { strIndex });
                Type typeDocMT_TSP_DateProp = oDocMT_TSP_DateProp.GetType();
                strValue = typeDocMT_TSP_DateProp.InvokeMember("Value",
                                           BindingFlags.Default |
                                           BindingFlags.GetProperty,
                                           null, oDocMT_TSP_DateProp,
                                           new object[] { }).ToString();

                // Set the MT_TSP_Date strindex to the text set in Project Date

                typeDocMT_TSP_DateProp.InvokeMember("Item",
                    BindingFlags.Default |
                    BindingFlags.SetProperty,
                    null, oDocCustomProps,
                    new object[] { strIndex, dateTimePicker1.Text });

                // Update all fields 

                // Update main body
                oDoc.Fields.Update();
               

                // Loop through all sections
                foreach (Section section in oDoc.Sections)
                {
                    //Get all Headers
                    HeadersFooters headers = section.Headers;

                    // Update fields in header
                    foreach (HeaderFooter header in headers)
                    {
                        Fields fields = header.Range.Fields;
                        foreach (Field field in fields)
                            field.Update();
                    }

                    //Get all footers
                    HeadersFooters footers = section.Footers;

                    // Update fields in footer
                    foreach (HeaderFooter footer in footers)
                    {
                        Fields fields = footer.Range.Fields;
                        foreach (Field field in fields)
                            field.Update();
                    }
                }

                // Update table of contents
                oDoc.TablesOfContents[1].Update();

                // Save and close word document
                
                oDoc.Save();
                oDoc.Close();

                // Quitting the work application so WINWORD.EXE does not persist in memory

                oWord.Quit();

                // Releasing COM Objects (Normal.dotm was being in use if this was not done)

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.FinalReleaseComObject(oDoc);

                    
             }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void exitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Really Quit?", "Exit", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                System.Windows.Forms.Application.Exit();
               
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Document Control Form Rev. 1");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            if (String.IsNullOrEmpty(outputfolderpath))
            {
                MessageBox.Show("Output Folder Path not specified");
                return;
            }

            if (MessageBox.Show("Did you fill the Document Properties", "Alert", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {

                button1_Click(null, null);
                UpdateWordProperties_Click(null, null);
                button2_Click(null, null);
                ButtonZIP_Click(null, null);

            }
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        
    }

}

