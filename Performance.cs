using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for Performance.
	/// </summary>
	public class Performance : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.NumericUpDown nudDelay;
        private System.Windows.Forms.ToolTip ttpBatchMail;
        private System.Windows.Forms.Label lblDelay;
        private System.Windows.Forms.ComboBox cboBCC;
        private System.Windows.Forms.ComboBox cboCC;
        private System.Windows.Forms.CheckBox chkGUID;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.LinkLabel lnkFolder;
        private System.Windows.Forms.CheckBox chkAttach;
        private System.Windows.Forms.NumericUpDown nudLoop;
        private System.Windows.Forms.TextBox txtFrom;
        private System.Windows.Forms.ComboBox cboTo;
        private System.Windows.Forms.GroupBox gbxOutLook;
        private System.Windows.Forms.LinkLabel lnkFile;
        private System.Windows.Forms.RadioButton rdoFileCase;
        private System.Windows.Forms.TextBox txtMailAddrFile;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Label lblLoop;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.ComboBox cboSMTP;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.Label lblSMTP;
        private System.Windows.Forms.RichTextBox richBox;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.LinkLabel lnkBCC;
        private System.Windows.Forms.LinkLabel lnkCC;
        private System.Windows.Forms.LinkLabel lnkTo;
        private System.Windows.Forms.LinkLabel lnkFrom;
        private System.Windows.Forms.GroupBox gbxDigiSafe;
        private System.Windows.Forms.RadioButton rdoMailCase;
        private System.ComponentModel.IContainer components;

		public Performance()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.nudDelay = new System.Windows.Forms.NumericUpDown();
            this.ttpBatchMail = new System.Windows.Forms.ToolTip(this.components);
            this.lblDelay = new System.Windows.Forms.Label();
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.chkGUID = new System.Windows.Forms.CheckBox();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.chkAttach = new System.Windows.Forms.CheckBox();
            this.nudLoop = new System.Windows.Forms.NumericUpDown();
            this.txtFrom = new System.Windows.Forms.TextBox();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.gbxOutLook = new System.Windows.Forms.GroupBox();
            this.lnkFile = new System.Windows.Forms.LinkLabel();
            this.rdoFileCase = new System.Windows.Forms.RadioButton();
            this.txtMailAddrFile = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.lblLoop = new System.Windows.Forms.Label();
            this.btnTest = new System.Windows.Forms.Button();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.cboSMTP = new System.Windows.Forms.ComboBox();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblSMTP = new System.Windows.Forms.Label();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.lnkFrom = new System.Windows.Forms.LinkLabel();
            this.gbxDigiSafe = new System.Windows.Forms.GroupBox();
            this.rdoMailCase = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).BeginInit();
            this.gbxOutLook.SuspendLayout();
            this.gbxDigiSafe.SuspendLayout();
            this.SuspendLayout();
            // 
            // nudDelay
            // 
            this.nudDelay.Location = new System.Drawing.Point(208, 208);
            this.nudDelay.Maximum = new System.Decimal(new int[] {
                                                                     5,
                                                                     0,
                                                                     0,
                                                                     0});
            this.nudDelay.Name = "nudDelay";
            this.nudDelay.Size = new System.Drawing.Size(44, 20);
            this.nudDelay.TabIndex = 105;
            this.nudDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpBatchMail.SetToolTip(this.nudDelay, "sec (0..5)");
            // 
            // lblDelay
            // 
            this.lblDelay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDelay.Location = new System.Drawing.Point(172, 212);
            this.lblDelay.Name = "lblDelay";
            this.lblDelay.Size = new System.Drawing.Size(36, 16);
            this.lblDelay.TabIndex = 104;
            this.lblDelay.Text = "Delay";
            this.ttpBatchMail.SetToolTip(this.lblDelay, "in sec (0..5)");
            // 
            // cboBCC
            // 
            this.cboBCC.Location = new System.Drawing.Point(68, 148);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(312, 21);
            this.cboBCC.TabIndex = 101;
            this.ttpBatchMail.SetToolTip(this.cboBCC, "BCC To");
            // 
            // cboCC
            // 
            this.cboCC.Location = new System.Drawing.Point(68, 124);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(312, 21);
            this.cboCC.TabIndex = 100;
            this.ttpBatchMail.SetToolTip(this.cboCC, "CC To");
            // 
            // chkGUID
            // 
            this.chkGUID.Location = new System.Drawing.Point(84, 212);
            this.chkGUID.Name = "chkGUID";
            this.chkGUID.Size = new System.Drawing.Size(52, 16);
            this.chkGUID.TabIndex = 98;
            this.chkGUID.Text = "GUID";
            this.ttpBatchMail.SetToolTip(this.chkGUID, "Include GUID");
            // 
            // txtFolder
            // 
            this.txtFolder.Enabled = false;
            this.txtFolder.Location = new System.Drawing.Point(80, 232);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(132, 20);
            this.txtFolder.TabIndex = 97;
            this.txtFolder.Text = "c:\\TestData";
            this.ttpBatchMail.SetToolTip(this.txtFolder, "Path/Folder for attachments");
            // 
            // lnkFolder
            // 
            this.lnkFolder.Enabled = false;
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.Location = new System.Drawing.Point(4, 232);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(48, 16);
            this.lnkFolder.TabIndex = 96;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder :";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpBatchMail.SetToolTip(this.lnkFolder, "Browse the attachment folder");
            // 
            // chkAttach
            // 
            this.chkAttach.Location = new System.Drawing.Point(4, 212);
            this.chkAttach.Name = "chkAttach";
            this.chkAttach.Size = new System.Drawing.Size(84, 16);
            this.chkAttach.TabIndex = 94;
            this.chkAttach.Text = "Attachment";
            this.ttpBatchMail.SetToolTip(this.chkAttach, "Include attachements");
            // 
            // nudLoop
            // 
            this.nudLoop.Location = new System.Drawing.Point(316, 208);
            this.nudLoop.Maximum = new System.Decimal(new int[] {
                                                                    999999,
                                                                    0,
                                                                    0,
                                                                    0});
            this.nudLoop.Name = "nudLoop";
            this.nudLoop.Size = new System.Drawing.Size(64, 20);
            this.nudLoop.TabIndex = 93;
            this.nudLoop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpBatchMail.SetToolTip(this.nudLoop, "0 .. 999,999");
            this.nudLoop.Value = new System.Decimal(new int[] {
                                                                  1,
                                                                  0,
                                                                  0,
                                                                  0});
            // 
            // txtFrom
            // 
            this.txtFrom.Location = new System.Drawing.Point(68, 76);
            this.txtFrom.Name = "txtFrom";
            this.txtFrom.Size = new System.Drawing.Size(312, 20);
            this.txtFrom.TabIndex = 83;
            this.txtFrom.Text = "txtFrom";
            this.ttpBatchMail.SetToolTip(this.txtFrom, "A file contain a list of addresses");
            // 
            // cboTo
            // 
            this.cboTo.Items.AddRange(new object[] {
                                                       ""});
            this.cboTo.Location = new System.Drawing.Point(68, 100);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(312, 21);
            this.cboTo.TabIndex = 99;
            this.cboTo.Text = "login0@company1.zantaz.com";
            this.ttpBatchMail.SetToolTip(this.cboTo, "mail to");
            // 
            // gbxOutLook
            // 
            this.gbxOutLook.Controls.Add(this.lnkFile);
            this.gbxOutLook.Controls.Add(this.rdoFileCase);
            this.gbxOutLook.Controls.Add(this.txtMailAddrFile);
            this.gbxOutLook.Location = new System.Drawing.Point(4, 8);
            this.gbxOutLook.Name = "gbxOutLook";
            this.gbxOutLook.Size = new System.Drawing.Size(380, 44);
            this.gbxOutLook.TabIndex = 103;
            this.gbxOutLook.TabStop = false;
            // 
            // lnkFile
            // 
            this.lnkFile.Enabled = false;
            this.lnkFile.Location = new System.Drawing.Point(32, 24);
            this.lnkFile.Name = "lnkFile";
            this.lnkFile.Size = new System.Drawing.Size(32, 16);
            this.lnkFile.TabIndex = 2;
            this.lnkFile.TabStop = true;
            this.lnkFile.Text = "File :";
            this.ttpBatchMail.SetToolTip(this.lnkFile, "Locate the address file");
            // 
            // rdoFileCase
            // 
            this.rdoFileCase.Location = new System.Drawing.Point(8, 0);
            this.rdoFileCase.Name = "rdoFileCase";
            this.rdoFileCase.Size = new System.Drawing.Size(92, 16);
            this.rdoFileCase.TabIndex = 1;
            this.rdoFileCase.Text = "File Case";
            this.ttpBatchMail.SetToolTip(this.rdoFileCase, "Read the data from file");
            // 
            // txtMailAddrFile
            // 
            this.txtMailAddrFile.Enabled = false;
            this.txtMailAddrFile.Location = new System.Drawing.Point(64, 20);
            this.txtMailAddrFile.Name = "txtMailAddrFile";
            this.txtMailAddrFile.Size = new System.Drawing.Size(312, 20);
            this.txtMailAddrFile.TabIndex = 52;
            this.txtMailAddrFile.Text = "mail address  file";
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(316, 232);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(64, 21);
            this.btnSend.TabIndex = 95;
            this.btnSend.Text = "Send";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // lblLoop
            // 
            this.lblLoop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblLoop.Location = new System.Drawing.Point(256, 212);
            this.lblLoop.Name = "lblLoop";
            this.lblLoop.Size = new System.Drawing.Size(56, 16);
            this.lblLoop.TabIndex = 92;
            this.lblLoop.Text = "# of Loop";
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(316, 256);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(64, 21);
            this.btnTest.TabIndex = 91;
            this.btnTest.Text = "Test";
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 13;
            this.cboPort.Location = new System.Drawing.Point(260, 256);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(44, 21);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 90;
            this.cboPort.Text = "25";
            // 
            // cboSMTP
            // 
            this.cboSMTP.ItemHeight = 13;
            this.cboSMTP.Items.AddRange(new object[] {
                                                         ""});
            this.cboSMTP.Location = new System.Drawing.Point(80, 256);
            this.cboSMTP.Name = "cboSMTP";
            this.cboSMTP.Size = new System.Drawing.Size(132, 21);
            this.cboSMTP.Sorted = true;
            this.cboSMTP.TabIndex = 89;
            this.cboSMTP.Text = "10.1.89.201";
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(216, 260);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(38, 16);
            this.lblPort.TabIndex = 88;
            this.lblPort.Text = "Port # ";
            // 
            // lblSMTP
            // 
            this.lblSMTP.Location = new System.Drawing.Point(4, 260);
            this.lblSMTP.Name = "lblSMTP";
            this.lblSMTP.Size = new System.Drawing.Size(72, 16);
            this.lblSMTP.TabIndex = 87;
            this.lblSMTP.Text = "SMTP Server";
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(4, 288);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(376, 132);
            this.richBox.TabIndex = 86;
            this.richBox.Text = "richBox";
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(4, 188);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 85;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(68, 184);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(312, 20);
            this.txtSubject.TabIndex = 84;
            this.txtSubject.Text = "txtSubject";
            // 
            // lnkBCC
            // 
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(28, 152);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(36, 20);
            this.lnkBCC.TabIndex = 82;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkCC
            // 
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(36, 128);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(28, 16);
            this.lnkCC.TabIndex = 81;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(16, 104);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(48, 16);
            this.lnkTo.TabIndex = 80;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkFrom
            // 
            this.lnkFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFrom.Location = new System.Drawing.Point(16, 80);
            this.lnkFrom.Name = "lnkFrom";
            this.lnkFrom.Size = new System.Drawing.Size(48, 16);
            this.lnkFrom.TabIndex = 79;
            this.lnkFrom.TabStop = true;
            this.lnkFrom.Text = "From :";
            this.lnkFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // gbxDigiSafe
            // 
            this.gbxDigiSafe.Controls.Add(this.rdoMailCase);
            this.gbxDigiSafe.Location = new System.Drawing.Point(4, 56);
            this.gbxDigiSafe.Name = "gbxDigiSafe";
            this.gbxDigiSafe.Size = new System.Drawing.Size(380, 120);
            this.gbxDigiSafe.TabIndex = 102;
            this.gbxDigiSafe.TabStop = false;
            // 
            // rdoMailCase
            // 
            this.rdoMailCase.Checked = true;
            this.rdoMailCase.Location = new System.Drawing.Point(8, 0);
            this.rdoMailCase.Name = "rdoMailCase";
            this.rdoMailCase.Size = new System.Drawing.Size(88, 16);
            this.rdoMailCase.TabIndex = 0;
            this.rdoMailCase.TabStop = true;
            this.rdoMailCase.Text = "Mail Case";
            this.ttpBatchMail.SetToolTip(this.rdoMailCase, "Modified the mail header");
            // 
            // Performance
            // 
            this.Controls.Add(this.lblDelay);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.chkGUID);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.lnkFolder);
            this.Controls.Add(this.chkAttach);
            this.Controls.Add(this.nudLoop);
            this.Controls.Add(this.txtFrom);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.gbxOutLook);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.lblLoop);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.cboSMTP);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblSMTP);
            this.Controls.Add(this.richBox);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lnkBCC);
            this.Controls.Add(this.lnkCC);
            this.Controls.Add(this.lnkTo);
            this.Controls.Add(this.lnkFrom);
            this.Controls.Add(this.gbxDigiSafe);
            this.Controls.Add(this.nudDelay);
            this.Name = "Performance";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).EndInit();
            this.gbxOutLook.ResumeLayout(false);
            this.gbxDigiSafe.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

        private void btnSend_Click(object sender, System.EventArgs e)
        {
        
        }
	}
}
