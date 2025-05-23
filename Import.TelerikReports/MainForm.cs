using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Xml;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Import;
using Stimulsoft.Base;
using System.Collections.Generic;

namespace Import.TelerikReports
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnBrowseRpx;
        private System.Windows.Forms.Button btnBrowseStimulReport;
        private System.Windows.Forms.TextBox tbRpxFile;
        private System.Windows.Forms.TextBox tbStimulReportFile;
        private GroupBox groupBox1;
        private GroupBox groupBoxInformation;
        private TreeView treeViewLog;
        private CheckBox checkBoxUsePrimitives;
        private CheckBox checkBoxSetLinked;
        private GroupBox groupBox2;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //

            #region Settings
            StiSettings.Load();
            StiFormSettings.Load(this);
            #endregion
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnBrowseRpx = new System.Windows.Forms.Button();
            this.btnBrowseStimulReport = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tbRpxFile = new System.Windows.Forms.TextBox();
            this.tbStimulReportFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBoxInformation = new System.Windows.Forms.GroupBox();
            this.treeViewLog = new System.Windows.Forms.TreeView();
            this.checkBoxUsePrimitives = new System.Windows.Forms.CheckBox();
            this.checkBoxSetLinked = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBoxInformation.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowseRpx
            // 
            this.btnBrowseRpx.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseRpx.Location = new System.Drawing.Point(643, 39);
            this.btnBrowseRpx.Name = "btnBrowseRpx";
            this.btnBrowseRpx.Size = new System.Drawing.Size(24, 20);
            this.btnBrowseRpx.TabIndex = 2;
            this.btnBrowseRpx.Text = "...";
            this.btnBrowseRpx.Click += new System.EventHandler(this.btnBrowseRpx_Click);
            // 
            // btnBrowseStimulReport
            // 
            this.btnBrowseStimulReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseStimulReport.Location = new System.Drawing.Point(643, 89);
            this.btnBrowseStimulReport.Name = "btnBrowseStimulReport";
            this.btnBrowseStimulReport.Size = new System.Drawing.Size(24, 20);
            this.btnBrowseStimulReport.TabIndex = 5;
            this.btnBrowseStimulReport.Text = "...";
            this.btnBrowseStimulReport.Click += new System.EventHandler(this.btnBrowseStimulReport_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileName = "doc1";
            // 
            // tbRpxFile
            // 
            this.tbRpxFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbRpxFile.Location = new System.Drawing.Point(9, 39);
            this.tbRpxFile.Name = "tbRpxFile";
            this.tbRpxFile.Size = new System.Drawing.Size(633, 20);
            this.tbRpxFile.TabIndex = 1;
            // 
            // tbStimulReportFile
            // 
            this.tbStimulReportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbStimulReportFile.Location = new System.Drawing.Point(9, 90);
            this.tbStimulReportFile.Name = "tbStimulReportFile";
            this.tbStimulReportFile.Size = new System.Drawing.Size(633, 20);
            this.tbStimulReportFile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(18, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(430, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Telerik Reports Template:";
            // 
            // btnConvert
            // 
            this.btnConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnConvert.Location = new System.Drawing.Point(517, 367);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 3;
            this.btnConvert.Text = "Convert";
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(605, 367);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(18, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(430, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Stimulsoft Reports Template:";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnBrowseRpx);
            this.groupBox1.Controls.Add(this.btnBrowseStimulReport);
            this.groupBox1.Controls.Add(this.tbRpxFile);
            this.groupBox1.Controls.Add(this.tbStimulReportFile);
            this.groupBox1.Location = new System.Drawing.Point(8, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(677, 126);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Path to Report";
            // 
            // groupBoxInformation
            // 
            this.groupBoxInformation.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxInformation.Controls.Add(this.treeViewLog);
            this.groupBoxInformation.Location = new System.Drawing.Point(8, 205);
            this.groupBoxInformation.Name = "groupBoxInformation";
            this.groupBoxInformation.Size = new System.Drawing.Size(677, 152);
            this.groupBoxInformation.TabIndex = 2;
            this.groupBoxInformation.TabStop = false;
            this.groupBoxInformation.Text = "Information";
            // 
            // treeViewLog
            // 
            this.treeViewLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.treeViewLog.Location = new System.Drawing.Point(9, 20);
            this.treeViewLog.Name = "treeViewLog";
            this.treeViewLog.Size = new System.Drawing.Size(658, 123);
            this.treeViewLog.TabIndex = 0;
            // 
            // checkBoxUsePrimitives
            // 
            this.checkBoxUsePrimitives.AutoSize = true;
            this.checkBoxUsePrimitives.Checked = true;
            this.checkBoxUsePrimitives.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxUsePrimitives.Enabled = false;
            this.checkBoxUsePrimitives.Location = new System.Drawing.Point(9, 19);
            this.checkBoxUsePrimitives.Name = "checkBoxUsePrimitives";
            this.checkBoxUsePrimitives.Size = new System.Drawing.Size(215, 17);
            this.checkBoxUsePrimitives.TabIndex = 0;
            this.checkBoxUsePrimitives.Text = "Use primitives instead of shapes for Line";
            this.checkBoxUsePrimitives.UseVisualStyleBackColor = true;
            // 
            // checkBoxSetLinked
            // 
            this.checkBoxSetLinked.AutoSize = true;
            this.checkBoxSetLinked.Checked = true;
            this.checkBoxSetLinked.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSetLinked.Location = new System.Drawing.Point(9, 39);
            this.checkBoxSetLinked.Name = "checkBoxSetLinked";
            this.checkBoxSetLinked.Size = new System.Drawing.Size(207, 17);
            this.checkBoxSetLinked.TabIndex = 1;
            this.checkBoxSetLinked.Text = "Set Linked property for all components";
            this.checkBoxSetLinked.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.checkBoxUsePrimitives);
            this.groupBox2.Controls.Add(this.checkBoxSetLinked);
            this.groupBox2.Location = new System.Drawing.Point(8, 136);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(677, 63);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Options";
            // 
            // Form1
            // 
            this.AcceptButton = this.btnConvert;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(692, 398);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBoxInformation);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnConvert);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Convert Telerik Reports to Stimulsoft Reports";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBoxInformation.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Form1());
        }

        private void btnBrowseRpx_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select Telerik Reports file";
            openFileDialog1.Filter = "Telerik Reports file (*.trdx,*.trdp)|*.trdx;*.trdp";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbStimulReportFile.Text = Path.Combine(Path.GetDirectoryName(openFileDialog1.FileName), Path.GetFileNameWithoutExtension(openFileDialog1.FileName) + ".mrt");
                tbRpxFile.Text = openFileDialog1.FileName;
            }
        }

        private void btnBrowseStimulReport_Click(object sender, System.EventArgs e)
        {
            saveFileDialog1.Title = "Save Stimulsoft Reports file";
            saveFileDialog1.Filter = "Stimulsoft Reports file (*.mrt)|*.mrt";
            saveFileDialog1.InitialDirectory = openFileDialog1.InitialDirectory;
            //saveFileDialog1.FileName = Path.GetFileName(openFileDialog1.FileName);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbStimulReportFile.Text = saveFileDialog1.FileName;
            }
        }

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnConvert_Click(object sender, System.EventArgs e)
        {
            if (!File.Exists(tbRpxFile.Text))
            {
                MessageBox.Show("Incorrect Telerik Reports template file name",
                    null, MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                StiTreeViewLog log = new StiTreeViewLog(this.treeViewLog);

                log.OpenLog("Initialize ...");
                Application.DoEvents();
                StiReport report = new StiReport();

                log.CurrentNode.Text += "OK";
                log.CloseLog();
                log.OpenLog("Open Telerik Reports template file ...");
                Application.DoEvents();
                var bytes = File.ReadAllBytes(tbRpxFile.Text);

                log.CurrentNode.Text += "OK";
                log.CloseLog();
                log.OpenLog("Converting ...");
                Application.DoEvents();

                StiTelerikReportsHelper helper = new StiTelerikReportsHelper();
                var errorList = new List<string>();
                helper.ProcessFile(bytes, report, errorList);

                if (checkBoxSetLinked.Checked)
                {
                    foreach (StiComponent component in report.GetComponents())
                    {
                        component.Linked = true;
                    }
                }

                log.CurrentNode.Text += "OK";
                log.CloseLog();
                log.OpenLog("Save Stimulsoft Reports template file ...");
                Application.DoEvents();
                report.Save(tbStimulReportFile.Text);

                log.CurrentNode.Text += "OK";

                #if Test
                log.OpenLog("Conversion errors:");
                foreach (string st in errorList)
                {
                    log.WriteNode(st);
                }
                log.CloseLog();

                MessageBox.Show("Conversion complete!");

                report.Design();
                #endif
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            #region Settings
            StiSettings.Load();
            StiFormSettings.Save(this);

            StiSettings.Save();
            #endregion
        }

    }
}
