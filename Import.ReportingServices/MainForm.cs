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
using Stimulsoft.Base;
using Stimulsoft.Report.Import;

namespace Import.Rdl
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
        private System.Windows.Forms.Button btnBrowseRDL;
        private System.Windows.Forms.Button btnBrowseStimulReport;
        private System.Windows.Forms.TextBox tbRdlFile;
        private System.Windows.Forms.TextBox tbStimulReportFile;
        private GroupBox groupBox1;
        private GroupBox groupBoxInformation;
        private TreeView treeViewLog;
        private CheckBox checkBoxConvertSyntaxToCSharp;
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
            this.btnBrowseRDL = new System.Windows.Forms.Button();
            this.btnBrowseStimulReport = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tbRdlFile = new System.Windows.Forms.TextBox();
            this.tbStimulReportFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBoxInformation = new System.Windows.Forms.GroupBox();
            this.treeViewLog = new System.Windows.Forms.TreeView();
            this.checkBoxConvertSyntaxToCSharp = new System.Windows.Forms.CheckBox();
            this.checkBoxSetLinked = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBoxInformation.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowseRDL
            // 
            this.btnBrowseRDL.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseRDL.Location = new System.Drawing.Point(614, 57);
            this.btnBrowseRDL.Name = "btnBrowseRDL";
            this.btnBrowseRDL.Size = new System.Drawing.Size(38, 29);
            this.btnBrowseRDL.TabIndex = 2;
            this.btnBrowseRDL.Text = "...";
            this.btnBrowseRDL.Click += new System.EventHandler(this.btnBrowseRDL_Click);
            // 
            // btnBrowseStimulReport
            // 
            this.btnBrowseStimulReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseStimulReport.Location = new System.Drawing.Point(614, 130);
            this.btnBrowseStimulReport.Name = "btnBrowseStimulReport";
            this.btnBrowseStimulReport.Size = new System.Drawing.Size(38, 29);
            this.btnBrowseStimulReport.TabIndex = 5;
            this.btnBrowseStimulReport.Text = "...";
            this.btnBrowseStimulReport.Click += new System.EventHandler(this.btnBrowseStimulReport_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileName = "doc1";
            // 
            // tbRdlFile
            // 
            this.tbRdlFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbRdlFile.Location = new System.Drawing.Point(14, 57);
            this.tbRdlFile.Name = "tbRdlFile";
            this.tbRdlFile.Size = new System.Drawing.Size(598, 26);
            this.tbRdlFile.TabIndex = 1;
            // 
            // tbStimulReportFile
            // 
            this.tbStimulReportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbStimulReportFile.Location = new System.Drawing.Point(14, 132);
            this.tbStimulReportFile.Name = "tbStimulReportFile";
            this.tbStimulReportFile.Size = new System.Drawing.Size(598, 26);
            this.tbStimulReportFile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(29, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(209, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "RDL Template:";
            // 
            // btnConvert
            // 
            this.btnConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnConvert.Location = new System.Drawing.Point(412, 352);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(120, 34);
            this.btnConvert.TabIndex = 3;
            this.btnConvert.Text = "Convert";
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(553, 352);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(120, 34);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(29, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(688, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "Stimulsoft Reports Template:";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnBrowseRDL);
            this.groupBox1.Controls.Add(this.btnBrowseStimulReport);
            this.groupBox1.Controls.Add(this.tbRdlFile);
            this.groupBox1.Controls.Add(this.tbStimulReportFile);
            this.groupBox1.Location = new System.Drawing.Point(13, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(668, 184);
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
            this.groupBoxInformation.Location = new System.Drawing.Point(13, 300);
            this.groupBoxInformation.Name = "groupBoxInformation";
            this.groupBoxInformation.Size = new System.Drawing.Size(668, 38);
            this.groupBoxInformation.TabIndex = 2;
            this.groupBoxInformation.TabStop = false;
            this.groupBoxInformation.Text = "Information";
            // 
            // treeViewLog
            // 
            this.treeViewLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeViewLog.Location = new System.Drawing.Point(14, 29);
            this.treeViewLog.Name = "treeViewLog";
            this.treeViewLog.Size = new System.Drawing.Size(638, 0);
            this.treeViewLog.TabIndex = 0;
            // 
            // checkBoxConvertSyntaxToCSharp
            // 
            this.checkBoxConvertSyntaxToCSharp.AutoSize = true;
            this.checkBoxConvertSyntaxToCSharp.Location = new System.Drawing.Point(14, 28);
            this.checkBoxConvertSyntaxToCSharp.Name = "checkBoxConvertSyntaxToCSharp";
            this.checkBoxConvertSyntaxToCSharp.Size = new System.Drawing.Size(208, 24);
            this.checkBoxConvertSyntaxToCSharp.TabIndex = 0;
            this.checkBoxConvertSyntaxToCSharp.Text = "Convert the syntax to C#";
            this.checkBoxConvertSyntaxToCSharp.UseVisualStyleBackColor = true;
            // 
            // checkBoxSetLinked
            // 
            this.checkBoxSetLinked.AutoSize = true;
            this.checkBoxSetLinked.Checked = true;
            this.checkBoxSetLinked.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSetLinked.Location = new System.Drawing.Point(14, 57);
            this.checkBoxSetLinked.Name = "checkBoxSetLinked";
            this.checkBoxSetLinked.Size = new System.Drawing.Size(307, 24);
            this.checkBoxSetLinked.TabIndex = 1;
            this.checkBoxSetLinked.Text = "Set Linked property for all components";
            this.checkBoxSetLinked.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.checkBoxConvertSyntaxToCSharp);
            this.groupBox2.Controls.Add(this.checkBoxSetLinked);
            this.groupBox2.Location = new System.Drawing.Point(13, 199);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(668, 92);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Options";
            // 
            // Form1
            // 
            this.AcceptButton = this.btnConvert;
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
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
            this.Text = "Convert RDL to Stimulsoft Reports";
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

        private void btnBrowseRDL_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select RDL file";
            openFileDialog1.Filter = "RDL file (*.rdl, *.rdlc)|*.rdl;*.rdlc";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbStimulReportFile.Text = Path.Combine(Path.GetDirectoryName(openFileDialog1.FileName), Path.GetFileNameWithoutExtension(openFileDialog1.FileName) + ".mrt");
                tbRdlFile.Text = openFileDialog1.FileName;
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
            if (!File.Exists(tbRdlFile.Text))
            {
                MessageBox.Show("Incorrect RDL template file name",
                    null, MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                StiTreeViewLog log = new StiTreeViewLog(this.treeViewLog);

                log.OpenLog("Initialize ...");
                Application.DoEvents();
                StiReport report = new StiReport();
                if (!string.IsNullOrWhiteSpace(tbStimulReportFile.Text)) report.ReportName = Path.GetFileNameWithoutExtension(tbStimulReportFile.Text);

                log.CurrentNode.Text += "OK";
                log.CloseLog();
                log.OpenLog("Open RDL template file ...");
                Application.DoEvents();

                byte[] bytes = File.ReadAllBytes(tbRdlFile.Text);

                XmlDocument doc = new XmlDocument();
                doc.Load(new MemoryStream(bytes));

                log.CurrentNode.Text += "OK";
                log.CloseLog();
                log.OpenLog("Converting ...");
                Application.DoEvents();

                StiReportingServicesHelper helper = new StiReportingServicesHelper() { ConvertSyntaxToCSharp = checkBoxConvertSyntaxToCSharp.Checked };
                ArrayList errorList = new ArrayList();
                helper.ProcessRootNode(doc.DocumentElement, report, errorList);

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
                foreach (string st in result.Errors)
                {
                    log.WriteNode(st);
                }
                log.CloseLog();
                #endif

                MessageBox.Show("Conversion complete!");

                #if Test
                result.Report.Design();
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
