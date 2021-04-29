using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;
using System.Web;
using System.Web.Mail;

namespace QATool
{
	/// <summary>
	/// Summary description for BatMailPage.
	/// </summary>
	public class BatMailPage : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.ComboBox cboPort;
		private System.Windows.Forms.ComboBox cboSMTP;
		private System.Windows.Forms.Label lblPort;
		private System.Windows.Forms.Label lblSMTP;
		private System.Windows.Forms.RichTextBox richBox;
		private System.Windows.Forms.Label lblSubject;
		private System.Windows.Forms.TextBox txtSubject;
		private System.Windows.Forms.TextBox txtFrom;
		private System.Windows.Forms.LinkLabel lnkBCC;
		private System.Windows.Forms.LinkLabel lnkCC;
		private System.Windows.Forms.LinkLabel lnkTo;
		private System.Windows.Forms.LinkLabel lnkFrom;
		private System.Windows.Forms.Label lblLoop;
		private System.Windows.Forms.CheckBox chkAttach;
		private System.Windows.Forms.Button btnSend;
		private System.Windows.Forms.ToolTip ttpBatchMail;
		private System.Windows.Forms.LinkLabel lnkFolder;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.NumericUpDown nudLoop;
		private System.ComponentModel.IContainer components;

		// custom declaration
		private String msgCaption = "Batch Mail Page";
		private System.Windows.Forms.CheckBox chkGUID;
		private System.Windows.Forms.ComboBox cboTo;
		private System.Windows.Forms.ComboBox cboCC;
		private System.Windows.Forms.ComboBox cboBCC; 

		private QATool.CommObj   commObj   = new CommObj();
		private QATool.AttachObj attachObj = null;
		private System.Windows.Forms.GroupBox gbxDigiSafe;
		private System.Windows.Forms.RadioButton rdoNormal;
		private System.Windows.Forms.GroupBox gbxOutLook;
		private System.Windows.Forms.RadioButton rdoOLCase;
		private System.Windows.Forms.LinkLabel lnkFile;
		private System.Windows.Forms.TextBox txtMailAddrFile;
		private System.Windows.Forms.NumericUpDown nudDelay;
		private System.Windows.Forms.Label lblDelay;
        private System.Windows.Forms.GroupBox gpbSMTP;
        private System.Windows.Forms.RadioButton rdoSMTP;
        private System.Windows.Forms.Label lblRcptTo;
        private System.Windows.Forms.ComboBox cboRcptTo;
        private System.Windows.Forms.ComboBox cboMailFrom;
        private System.Windows.Forms.LinkLabel lnkFileStream;
        private System.Windows.Forms.TextBox txtInputFile;
        private System.Windows.Forms.Label lblMailFrom;		
        private System.Windows.Forms.Label lblThread;
        private System.Windows.Forms.NumericUpDown nudThread;
        private System.Windows.Forms.Button btnAbort;

        private Thread smtpMailThread;
        private Thread fileMailThread;
        private Thread sendMailThread;

//		private static int m_idxAttach = 0; // index of a file list
//		private static int m_numFile   = 0; // number of test data

		public BatMailPage()
		{
            Debug.WriteLine("BatMailPage.cs - Initialize BatMailPage Object");
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
			ttpBatchMail.SetToolTip( lnkFrom,   "Point to file that contain addresses" );
			ttpBatchMail.SetToolTip( lnkTo,     "Load the address book" );
			ttpBatchMail.SetToolTip( lnkCC,     "Load the address book" );
			ttpBatchMail.SetToolTip( lnkBCC,    "Load the address book" );
			ttpBatchMail.SetToolTip( lnkFolder, "Specific attachment data location" );
			ttpBatchMail.SetToolTip( lblLoop,   "Repeat count - 1 .. 9999" );
			ttpBatchMail.SetToolTip( txtFrom,   "Add validation");
			ttpBatchMail.SetToolTip( cboTo,     "Add validation");
			ttpBatchMail.SetToolTip( cboCC,     "Add validation");
			ttpBatchMail.SetToolTip( cboBCC,    "Add validation");
			ttpBatchMail.SetToolTip( txtSubject,"Add validation");

			commObj.InitComboBoxItem( cboTo, "[To Address]" );
			commObj.InitComboBoxItem( cboCC, "[CC Address]" );
			commObj.InitComboBoxItem( cboBCC, "[BCC Address]" );
			commObj.InitComboBoxItem( cboSMTP, "[SMTP IP]" );
			commObj.InitComboBoxItem( cboPort, "[Port]" );	
		    commObj.InitComboBoxItem( cboRcptTo, "[To Address]" );
            commObj.InitComboBoxItem( cboMailFrom, "[To Address]" );
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
            if( sendMailThread != null && sendMailThread.IsAlive )
                this.KillSendMailThread();

            if( smtpMailThread != null && smtpMailThread.IsAlive )
                this.KillSmtpMailThread();

            if( fileMailThread != null && fileMailThread.IsAlive )
                this.KillFileMailThread();

			if( disposing )
			{
                Debug.WriteLine( "BatMailPage.cs - Deposing BatMailPage Object");
                commObj.LogToFile("BatMailPage.cs - Deposing BatMailPage Object");

				if(components != null)
				{                    					
					components.Dispose();
                    commObj.LogToFile("\tDispose BatMailPage component");
                    Debug.WriteLine("\t Dispose component");
				}
			}
			base.Dispose( disposing );
		}//end of Dispose

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.btnTest = new System.Windows.Forms.Button();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.cboSMTP = new System.Windows.Forms.ComboBox();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblSMTP = new System.Windows.Forms.Label();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.txtFrom = new System.Windows.Forms.TextBox();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.lnkFrom = new System.Windows.Forms.LinkLabel();
            this.lblLoop = new System.Windows.Forms.Label();
            this.nudLoop = new System.Windows.Forms.NumericUpDown();
            this.chkAttach = new System.Windows.Forms.CheckBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.ttpBatchMail = new System.Windows.Forms.ToolTip(this.components);
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.chkGUID = new System.Windows.Forms.CheckBox();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.rdoNormal = new System.Windows.Forms.RadioButton();
            this.lnkFile = new System.Windows.Forms.LinkLabel();
            this.rdoOLCase = new System.Windows.Forms.RadioButton();
            this.nudDelay = new System.Windows.Forms.NumericUpDown();
            this.lblDelay = new System.Windows.Forms.Label();
            this.rdoSMTP = new System.Windows.Forms.RadioButton();
            this.lnkFileStream = new System.Windows.Forms.LinkLabel();
            this.txtInputFile = new System.Windows.Forms.TextBox();
            this.cboRcptTo = new System.Windows.Forms.ComboBox();
            this.lblThread = new System.Windows.Forms.Label();
            this.nudThread = new System.Windows.Forms.NumericUpDown();
            this.cboMailFrom = new System.Windows.Forms.ComboBox();
            this.lblRcptTo = new System.Windows.Forms.Label();
            this.btnAbort = new System.Windows.Forms.Button();
            this.gbxDigiSafe = new System.Windows.Forms.GroupBox();
            this.gbxOutLook = new System.Windows.Forms.GroupBox();
            this.txtMailAddrFile = new System.Windows.Forms.TextBox();
            this.gpbSMTP = new System.Windows.Forms.GroupBox();
            this.lblMailFrom = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudThread)).BeginInit();
            this.gbxDigiSafe.SuspendLayout();
            this.gbxOutLook.SuspendLayout();
            this.gpbSMTP.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(280, 328);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(48, 21);
            this.btnTest.TabIndex = 35;
            this.btnTest.Text = "Test";
            this.ttpBatchMail.SetToolTip(this.btnTest, "Test SMTP connection");
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 13;
            this.cboPort.Location = new System.Drawing.Point(224, 328);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(52, 21);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 34;
            this.cboPort.Text = "25";
            this.ttpBatchMail.SetToolTip(this.cboPort, "port number");
            // 
            // cboSMTP
            // 
            this.cboSMTP.ItemHeight = 13;
            this.cboSMTP.Items.AddRange(new object[] {
                                                         ""});
            this.cboSMTP.Location = new System.Drawing.Point(44, 328);
            this.cboSMTP.Name = "cboSMTP";
            this.cboSMTP.Size = new System.Drawing.Size(144, 21);
            this.cboSMTP.Sorted = true;
            this.cboSMTP.TabIndex = 33;
            this.cboSMTP.Text = "10.1.89.201";
            this.ttpBatchMail.SetToolTip(this.cboSMTP, "Server name or IP");
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(188, 332);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(28, 16);
            this.lblPort.TabIndex = 32;
            this.lblPort.Text = "Port";
            this.ttpBatchMail.SetToolTip(this.lblPort, "SMTP port number");
            // 
            // lblSMTP
            // 
            this.lblSMTP.Location = new System.Drawing.Point(0, 332);
            this.lblSMTP.Name = "lblSMTP";
            this.lblSMTP.Size = new System.Drawing.Size(36, 16);
            this.lblSMTP.TabIndex = 31;
            this.lblSMTP.Text = "SMTP";
            this.ttpBatchMail.SetToolTip(this.lblSMTP, "SMTP Server name or IP");
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(4, 356);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(380, 64);
            this.richBox.TabIndex = 30;
            this.richBox.Text = "richBox";
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(60, 260);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 29;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(116, 256);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(264, 20);
            this.txtSubject.TabIndex = 28;
            this.txtSubject.Text = "txtSubject";
            this.ttpBatchMail.SetToolTip(this.txtSubject, "Empty it if only want to show GUID");
            // 
            // txtFrom
            // 
            this.txtFrom.Location = new System.Drawing.Point(68, 60);
            this.txtFrom.Name = "txtFrom";
            this.txtFrom.Size = new System.Drawing.Size(312, 20);
            this.txtFrom.TabIndex = 24;
            this.txtFrom.Text = "txtFrom";
            this.ttpBatchMail.SetToolTip(this.txtFrom, "A file contain a list of addresses");
            // 
            // lnkBCC
            // 
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(28, 136);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(36, 20);
            this.lnkBCC.TabIndex = 23;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkBCC.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkBCC_LinkClicked);
            // 
            // lnkCC
            // 
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(36, 112);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(28, 16);
            this.lnkCC.TabIndex = 22;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkCC.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCC_LinkClicked);
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(16, 88);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(48, 16);
            this.lnkTo.TabIndex = 21;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkTo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkTo_LinkClicked);
            // 
            // lnkFrom
            // 
            this.lnkFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFrom.Location = new System.Drawing.Point(16, 64);
            this.lnkFrom.Name = "lnkFrom";
            this.lnkFrom.Size = new System.Drawing.Size(48, 16);
            this.lnkFrom.TabIndex = 20;
            this.lnkFrom.TabStop = true;
            this.lnkFrom.Text = "From :";
            this.lnkFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkFrom.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFrom_LinkClicked);
            // 
            // lblLoop
            // 
            this.lblLoop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblLoop.Location = new System.Drawing.Point(0, 308);
            this.lblLoop.Name = "lblLoop";
            this.lblLoop.Size = new System.Drawing.Size(40, 16);
            this.lblLoop.TabIndex = 36;
            this.lblLoop.Text = "# Loop";
            this.ttpBatchMail.SetToolTip(this.lblLoop, "0 .. 999,999");
            // 
            // nudLoop
            // 
            this.nudLoop.Location = new System.Drawing.Point(44, 304);
            this.nudLoop.Maximum = new System.Decimal(new int[] {
                                                                    999999,
                                                                    0,
                                                                    0,
                                                                    0});
            this.nudLoop.Name = "nudLoop";
            this.nudLoop.Size = new System.Drawing.Size(76, 20);
            this.nudLoop.TabIndex = 39;
            this.nudLoop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpBatchMail.SetToolTip(this.nudLoop, "0 .. 999,999");
            this.nudLoop.Value = new System.Decimal(new int[] {
                                                                  1,
                                                                  0,
                                                                  0,
                                                                  0});
            // 
            // chkAttach
            // 
            this.chkAttach.Location = new System.Drawing.Point(4, 284);
            this.chkAttach.Name = "chkAttach";
            this.chkAttach.Size = new System.Drawing.Size(56, 16);
            this.chkAttach.TabIndex = 40;
            this.chkAttach.Text = "Attach";
            this.ttpBatchMail.SetToolTip(this.chkAttach, "Include attachements");
            this.chkAttach.CheckedChanged += new System.EventHandler(this.chkAttach_CheckedChanged);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(332, 304);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(48, 21);
            this.btnSend.TabIndex = 41;
            this.btnSend.Text = "Send";
            this.ttpBatchMail.SetToolTip(this.btnSend, "Initial Sending thread");
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // lnkFolder
            // 
            this.lnkFolder.Enabled = false;
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkFolder.Location = new System.Drawing.Point(60, 284);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(56, 16);
            this.lnkFolder.TabIndex = 42;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder:";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpBatchMail.SetToolTip(this.lnkFolder, "Browse the attachment folder");
            this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
            // 
            // txtFolder
            // 
            this.txtFolder.Enabled = false;
            this.txtFolder.Location = new System.Drawing.Point(116, 280);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(264, 20);
            this.txtFolder.TabIndex = 43;
            this.txtFolder.Text = "c:\\TestData";
            this.ttpBatchMail.SetToolTip(this.txtFolder, "Path/Folder for attachments");
            // 
            // chkGUID
            // 
            this.chkGUID.Location = new System.Drawing.Point(4, 260);
            this.chkGUID.Name = "chkGUID";
            this.chkGUID.Size = new System.Drawing.Size(52, 16);
            this.chkGUID.TabIndex = 46;
            this.chkGUID.Text = "GUID";
            this.ttpBatchMail.SetToolTip(this.chkGUID, "Include GUID");
            // 
            // cboTo
            // 
            this.cboTo.Items.AddRange(new object[] {
                                                       ""});
            this.cboTo.Location = new System.Drawing.Point(68, 84);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(312, 21);
            this.cboTo.TabIndex = 47;
            this.cboTo.Text = "login0@company1.zantaz.com";
            this.ttpBatchMail.SetToolTip(this.cboTo, "mail to");
            // 
            // cboCC
            // 
            this.cboCC.Location = new System.Drawing.Point(68, 108);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(312, 21);
            this.cboCC.TabIndex = 48;
            this.ttpBatchMail.SetToolTip(this.cboCC, "CC To");
            // 
            // cboBCC
            // 
            this.cboBCC.Location = new System.Drawing.Point(68, 132);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(312, 21);
            this.cboBCC.TabIndex = 49;
            this.ttpBatchMail.SetToolTip(this.cboBCC, "BCC To");
            // 
            // rdoNormal
            // 
            this.rdoNormal.Checked = true;
            this.rdoNormal.Location = new System.Drawing.Point(8, 0);
            this.rdoNormal.Name = "rdoNormal";
            this.rdoNormal.Size = new System.Drawing.Size(88, 16);
            this.rdoNormal.TabIndex = 0;
            this.rdoNormal.TabStop = true;
            this.rdoNormal.Text = "Normal Case";
            this.ttpBatchMail.SetToolTip(this.rdoNormal, "Send mail by MS API");
            this.rdoNormal.Click += new System.EventHandler(this.rdoNormal_Click);
            // 
            // lnkFile
            // 
            this.lnkFile.Enabled = false;
            this.lnkFile.Location = new System.Drawing.Point(32, 20);
            this.lnkFile.Name = "lnkFile";
            this.lnkFile.Size = new System.Drawing.Size(32, 16);
            this.lnkFile.TabIndex = 2;
            this.lnkFile.TabStop = true;
            this.lnkFile.Text = "File :";
            this.ttpBatchMail.SetToolTip(this.lnkFile, "Locate the address file");
            this.lnkFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFile_LinkClicked);
            // 
            // rdoOLCase
            // 
            this.rdoOLCase.Location = new System.Drawing.Point(8, 0);
            this.rdoOLCase.Name = "rdoOLCase";
            this.rdoOLCase.Size = new System.Drawing.Size(92, 16);
            this.rdoOLCase.TabIndex = 0;
            this.rdoOLCase.Text = "Outlook/Notes";
            this.ttpBatchMail.SetToolTip(this.rdoOLCase, "Simulate sending mails between outlook accounts");
            this.rdoOLCase.Click += new System.EventHandler(this.rdoOLCase_Click);
            // 
            // nudDelay
            // 
            this.nudDelay.Location = new System.Drawing.Point(276, 304);
            this.nudDelay.Maximum = new System.Decimal(new int[] {
                                                                     5,
                                                                     0,
                                                                     0,
                                                                     0});
            this.nudDelay.Name = "nudDelay";
            this.nudDelay.Size = new System.Drawing.Size(52, 20);
            this.nudDelay.TabIndex = 78;
            this.nudDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpBatchMail.SetToolTip(this.nudDelay, "sec (0..5)");
            // 
            // lblDelay
            // 
            this.lblDelay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDelay.Location = new System.Drawing.Point(240, 308);
            this.lblDelay.Name = "lblDelay";
            this.lblDelay.Size = new System.Drawing.Size(36, 16);
            this.lblDelay.TabIndex = 77;
            this.lblDelay.Text = "Delay";
            this.ttpBatchMail.SetToolTip(this.lblDelay, "in sec (0..5)");
            // 
            // rdoSMTP
            // 
            this.rdoSMTP.Location = new System.Drawing.Point(8, 0);
            this.rdoSMTP.Name = "rdoSMTP";
            this.rdoSMTP.Size = new System.Drawing.Size(84, 16);
            this.rdoSMTP.TabIndex = 1;
            this.rdoSMTP.Text = "SMTP Case";
            this.ttpBatchMail.SetToolTip(this.rdoSMTP, "Stream a file to an SMTP socket");
            this.rdoSMTP.Click += new System.EventHandler(this.rdoSMTP_Click);
            // 
            // lnkFileStream
            // 
            this.lnkFileStream.Enabled = false;
            this.lnkFileStream.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFileStream.Location = new System.Drawing.Point(4, 68);
            this.lnkFileStream.Name = "lnkFileStream";
            this.lnkFileStream.Size = new System.Drawing.Size(56, 16);
            this.lnkFileStream.TabIndex = 5;
            this.lnkFileStream.TabStop = true;
            this.lnkFileStream.Text = "Input File";
            this.ttpBatchMail.SetToolTip(this.lnkFileStream, "Stream this file");
            this.lnkFileStream.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFileStream_LinkClicked);
            // 
            // txtInputFile
            // 
            this.txtInputFile.Enabled = false;
            this.txtInputFile.Location = new System.Drawing.Point(64, 64);
            this.txtInputFile.Name = "txtInputFile";
            this.txtInputFile.Size = new System.Drawing.Size(312, 20);
            this.txtInputFile.TabIndex = 6;
            this.txtInputFile.Text = "";
            this.ttpBatchMail.SetToolTip(this.txtInputFile, "file that stream into socket");
            // 
            // cboRcptTo
            // 
            this.cboRcptTo.Enabled = false;
            this.cboRcptTo.Location = new System.Drawing.Point(64, 15);
            this.cboRcptTo.Name = "cboRcptTo";
            this.cboRcptTo.Size = new System.Drawing.Size(312, 21);
            this.cboRcptTo.TabIndex = 2;
            this.ttpBatchMail.SetToolTip(this.cboRcptTo, "Type in or select from the pull down");
            // 
            // lblThread
            // 
            this.lblThread.Location = new System.Drawing.Point(124, 308);
            this.lblThread.Name = "lblThread";
            this.lblThread.Size = new System.Drawing.Size(52, 16);
            this.lblThread.TabIndex = 80;
            this.lblThread.Text = "# Thread";
            this.ttpBatchMail.SetToolTip(this.lblThread, "UNDER CONSTRUCTION");
            // 
            // nudThread
            // 
            this.nudThread.Location = new System.Drawing.Point(176, 304);
            this.nudThread.Name = "nudThread";
            this.nudThread.Size = new System.Drawing.Size(52, 20);
            this.nudThread.TabIndex = 81;
            this.nudThread.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpBatchMail.SetToolTip(this.nudThread, "1 .. 10");
            this.nudThread.Value = new System.Decimal(new int[] {
                                                                    1,
                                                                    0,
                                                                    0,
                                                                    0});
            // 
            // cboMailFrom
            // 
            this.cboMailFrom.Enabled = false;
            this.cboMailFrom.Location = new System.Drawing.Point(64, 40);
            this.cboMailFrom.Name = "cboMailFrom";
            this.cboMailFrom.Size = new System.Drawing.Size(312, 21);
            this.cboMailFrom.TabIndex = 4;
            this.cboMailFrom.Text = "combo@combo.com";
            this.ttpBatchMail.SetToolTip(this.cboMailFrom, "Type in or select from pull down box");
            // 
            // lblRcptTo
            // 
            this.lblRcptTo.Enabled = false;
            this.lblRcptTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblRcptTo.Location = new System.Drawing.Point(12, 20);
            this.lblRcptTo.Name = "lblRcptTo";
            this.lblRcptTo.Size = new System.Drawing.Size(48, 16);
            this.lblRcptTo.TabIndex = 1;
            this.lblRcptTo.Text = "Rcpt To";
            this.ttpBatchMail.SetToolTip(this.lblRcptTo, "Rcpt To");
            // 
            // btnAbort
            // 
            this.btnAbort.Location = new System.Drawing.Point(332, 328);
            this.btnAbort.Name = "btnAbort";
            this.btnAbort.Size = new System.Drawing.Size(48, 21);
            this.btnAbort.TabIndex = 82;
            this.btnAbort.Text = "Abort";
            this.ttpBatchMail.SetToolTip(this.btnAbort, "Kill the Sending thread... patient");
            this.btnAbort.Click += new System.EventHandler(this.btnAbort_Click);
            // 
            // gbxDigiSafe
            // 
            this.gbxDigiSafe.Controls.Add(this.rdoNormal);
            this.gbxDigiSafe.Location = new System.Drawing.Point(4, 40);
            this.gbxDigiSafe.Name = "gbxDigiSafe";
            this.gbxDigiSafe.Size = new System.Drawing.Size(380, 120);
            this.gbxDigiSafe.TabIndex = 50;
            this.gbxDigiSafe.TabStop = false;
            // 
            // gbxOutLook
            // 
            this.gbxOutLook.Controls.Add(this.lnkFile);
            this.gbxOutLook.Controls.Add(this.rdoOLCase);
            this.gbxOutLook.Controls.Add(this.txtMailAddrFile);
            this.gbxOutLook.Location = new System.Drawing.Point(4, 0);
            this.gbxOutLook.Name = "gbxOutLook";
            this.gbxOutLook.Size = new System.Drawing.Size(380, 40);
            this.gbxOutLook.TabIndex = 51;
            this.gbxOutLook.TabStop = false;
            // 
            // txtMailAddrFile
            // 
            this.txtMailAddrFile.Enabled = false;
            this.txtMailAddrFile.Location = new System.Drawing.Point(64, 16);
            this.txtMailAddrFile.Name = "txtMailAddrFile";
            this.txtMailAddrFile.Size = new System.Drawing.Size(312, 20);
            this.txtMailAddrFile.TabIndex = 52;
            this.txtMailAddrFile.Text = "mail address  file";
            // 
            // gpbSMTP
            // 
            this.gpbSMTP.Controls.Add(this.txtInputFile);
            this.gpbSMTP.Controls.Add(this.lnkFileStream);
            this.gpbSMTP.Controls.Add(this.cboMailFrom);
            this.gpbSMTP.Controls.Add(this.lblMailFrom);
            this.gpbSMTP.Controls.Add(this.cboRcptTo);
            this.gpbSMTP.Controls.Add(this.lblRcptTo);
            this.gpbSMTP.Controls.Add(this.rdoSMTP);
            this.gpbSMTP.Location = new System.Drawing.Point(4, 164);
            this.gpbSMTP.Name = "gpbSMTP";
            this.gpbSMTP.Size = new System.Drawing.Size(380, 88);
            this.gpbSMTP.TabIndex = 79;
            this.gpbSMTP.TabStop = false;
            // 
            // lblMailFrom
            // 
            this.lblMailFrom.Enabled = false;
            this.lblMailFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblMailFrom.Location = new System.Drawing.Point(4, 44);
            this.lblMailFrom.Name = "lblMailFrom";
            this.lblMailFrom.Size = new System.Drawing.Size(56, 16);
            this.lblMailFrom.TabIndex = 3;
            this.lblMailFrom.Text = "Mail From";
            // 
            // BatMailPage
            // 
            this.Controls.Add(this.btnAbort);
            this.Controls.Add(this.nudThread);
            this.Controls.Add(this.lblThread);
            this.Controls.Add(this.gpbSMTP);
            this.Controls.Add(this.nudDelay);
            this.Controls.Add(this.lblDelay);
            this.Controls.Add(this.gbxOutLook);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.chkGUID);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.chkAttach);
            this.Controls.Add(this.nudLoop);
            this.Controls.Add(this.lblLoop);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.cboSMTP);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblSMTP);
            this.Controls.Add(this.richBox);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.txtFrom);
            this.Controls.Add(this.lnkBCC);
            this.Controls.Add(this.lnkCC);
            this.Controls.Add(this.lnkTo);
            this.Controls.Add(this.lnkFrom);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.gbxDigiSafe);
            this.Controls.Add(this.lnkFolder);
            this.Name = "BatMailPage";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudThread)).EndInit();
            this.gbxDigiSafe.ResumeLayout(false);
            this.gbxOutLook.ResumeLayout(false);
            this.gpbSMTP.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

		private void chkAttach_CheckedChanged(object sender, System.EventArgs e)
		{
			if( chkAttach.Checked )
			{
				lnkFolder.Enabled = true;
				txtFolder.Enabled = true;
			}
			else
			{
				lnkFolder.Enabled = false;
				txtFolder.Enabled = false;
			}//end of else - disable
		}// end of chkAttach_CheckedChanged

		private void lnkFrom_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkFrom_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtFrom.Text = ofDlg.FileName;
			}//end of if		
		}// end of lnkFrom_LinkClicked

		private void lnkTo_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkTo_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				commObj.LoadComboBoxItem( cboTo, ofDlg.FileName );
			}//end of if		

		}// end of lnkTo_LinkClicked

		private void lnkCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkTo_LinkClicked" );
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				commObj.LoadComboBoxItem( cboCC, ofDlg.FileName );
			}//end of if		

		}// end of lnkCC_LinkClicked

		private void lnkBCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkTo_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				commObj.LoadComboBoxItem( cboBCC, ofDlg.FileName );
			}//end of if		

		}// end of lnkBCC_LinkClicked

		private void btnTest_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkTest_LinkClicked" );
			this.Cursor = Cursors.WaitCursor;
			richBox.Text = commObj.TestSMTPConnection(cboSMTP.Text,cboPort.Text)?"Connection OK":"Connection FAIL";
			this.Cursor = Cursors.Default;		
		}// end of btnTest_Click

		/// <summary>
		/// Generate a thread to send a batch mail....
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSend_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine("BatMailPage.cs - btnSend_Click");

			if( chkAttach.Checked )
			{
				attachObj = new AttachObj( txtFolder.Text );
//				DirectoryInfo di = new DirectoryInfo(txtFolder.Text); // attachment folder
//				FileInfo[] lstFiles = di.GetFiles();
//				m_numFile = lstFiles.Length;
			}// end of if - attachment check

            if( rdoNormal.Checked )
            {
                sendMailThread = new Thread( new ThreadStart(this.Thd_SendGuidMail) );
                sendMailThread.Name = "sendMailThread";
                sendMailThread.Start();
                commObj.LogToFile( "Thread.log", "++" + sendMailThread.Name + " Start ++");

            }//end of if - do the normal mailing
            else
                if( rdoOLCase.Checked )
                {
//                    AutoSendMail();

                    fileMailThread = new Thread( new ThreadStart(this.Thd_SendFileMail) );
                    fileMailThread.Name = "fileMailThread";
                    fileMailThread.Start();
                    commObj.LogToFile( "Thread.log", "++ fileMailThread Start ++");                    
                }//end of if - send to outlook base on a file 
            else
                if( rdoSMTP.Checked )
                {
                    smtpMailThread = new Thread( new ThreadStart(this.Thd_SendSmtpMail) );
                    smtpMailThread.Name = "smtpMailThread";
                    smtpMailThread.Start();
                    commObj.LogToFile( "Thread.log", "++ smtpMailThread Start ++");    
                }// end of if - stream file into socket

		}//end of btnSend_Click

		private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkFolder_LinkClicked" );
			FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
            if( txtFolder.Text != null )
                fbDlg.SelectedPath = txtFolder.Text;  // set the default folder

			if( fbDlg.ShowDialog() == DialogResult.OK )
			{
				txtFolder.Text = fbDlg.SelectedPath;
			}
		}//end of lnkFolder_LinkClicked

        public void HandleSmtpMail( String inStr )
        {
            Trace.WriteLine( "BatMailPage.cs - HandleSmtpMail" );

            try
            {
                QATool.SMTPSender smtpSender = new SMTPSender();
                smtpSender.mailFrom    = cboMailFrom.Text;
                smtpSender.mailTo      = cboRcptTo.Text;

                smtpSender.smtpServer  = cboSMTP.Text;
                smtpSender.smtpPortNum = cboPort.Text;

                smtpSender.SmtpSend( txtInputFile.Text );
            }//end of try
            catch( QATool.SMTPException ex )
            {
                Trace.WriteLine("\tSMTPClientSender() Exception: " + ex.SmtpMessage.ToString() );
                MessageBox.Show( ex.SmtpMessage.ToString(), msgCaption );
            }//end of catch
        }//end of HandleSmtpMail

		/// <summary>
		/// Constructing the mail based on the GUI setting and then send the mail
		/// Check - inlcude GUID in subject line and body
		/// Log the GUID into a text file for future searching
		/// Check - include attachment 
		/// </summary>
		public void HandleSendMail( String inStr )
		{
			Trace.WriteLine( "BatMailPage.cs - HandleSendMail" );

            int    counter = 0; // mail sent counter
			string strGUID = "";
			string strFrom;

            StreamReader sr = null;
			try
			{
				sr = new StreamReader( txtFrom.Text ); // address book - put here for exception catch
				while( (strFrom = sr.ReadLine()) != null ) // file name from txtFrom field
				{
                    counter++;
					Debug.WriteLine( "\t - HandleSendMail - inside while loop : " + counter.ToString() );

                    // May initial GUID here -> strGUID = "";

					richBox.Text = "Sent Mail: " + counter.ToString() + "\r\n- Read line : " + strFrom;
					if( chkGUID.Checked )
					{
						strGUID = System.Guid.NewGuid().ToString();
						txtSubject.Text = inStr + counter.ToString() + " " + strGUID;
						commObj.LogGUID( "GUID.LOG", strGUID );
					}//end of if - GUID			

					MailMessage mailMsg = new MailMessage();

					mailMsg.From	= strFrom;		// single line from file
					mailMsg.To		= cboTo.Text;	// single line from GUI
					mailMsg.Cc		= cboCC.Text;	// user input, may != To
					mailMsg.Bcc		= cboBCC.Text;	// user input, may != To
					mailMsg.Subject = txtSubject.Text;
					mailMsg.Body	= richBox.Text
						+ "\nFrom: " + strFrom		// start from here change
						+ "\n TO:  " + cboTo.Text
						+ "\n CC:  " + cboCC.Text
						+ "\n BCC: " + cboBCC.Text
						+ "\n Body_Subject: " + txtSubject.Text
						+ "\n" + DateTime.Now
                        + "\n Mail Counter == " + counter.ToString();

					if( chkAttach.Checked )
					{									
						if( attachObj.idxAttach == attachObj.numFile )
							attachObj.idxAttach = 0; //reset

						Debug.WriteLine( attachObj.idxAttach, "\t - idxAttach" );
						Debug.WriteLine( attachObj.attchFullName, "\t - filename" );
						mailMsg.Attachments.Add( new MailAttachment( attachObj.attchFullName, MailEncoding.Base64 ) );						

						mailMsg.Body += "\r\nAttach file index = " + attachObj.idxAttach
									  + "\r\nAttach file name = " + attachObj.attchFullName;

						attachObj.idxAttach++; // point to next file
					}//end of if

					try
					{
						Debug.WriteLine("  +  HandleSendMail - send mail la");
						richBox.Text += "\r\n+ Do the send mail";
                        SmtpMail.Send( mailMsg );
					}//end of try
					catch( System.Web.HttpException ex )
					{
						Debug.WriteLine(ex.Message.ToString());
						commObj.LogGUID( "GUID.LOG", "Exception occur" + strGUID.ToString() );
//						MessageBox.Show(ex.Message.ToString(), msgCaption);
					}// end of catch

					richBox.Text += "\r\n+ Batch Mails Sent Info " + mailMsg.Body; // display in rich Box
					commObj.LogToFile( richBox.Text ); // save into log

					Thread.Sleep( (int)nudDelay.Value * 1000 );
				}//end of while
			}//end of try - IOException
			catch( Exception ex )
			{
				Debug.WriteLine(ex.Message.ToString());
				MessageBox.Show(ex.Message.ToString(), msgCaption);
                commObj.LogToFile(ex.Message + "\n" + ex.StackTrace);
			}//end of catch - IOException	
			finally
			{
				if( sr != null )
				{
					Trace.WriteLine("Finally - close the Stream Reader");
                    commObj.LogToFile("Finally - close the Stream Reader");
					sr.Close();
				}//end of if
			}
		}//end of HandleSendMail

        public void Thd_SendSmtpMail()
        {
            Trace.WriteLine( "BatMailPage.cs - Thd_SendSmtpMail" );
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;
            DateTime startTime = DateTime.Now;

            for( int j = 0; j < nudLoop.Value; j++ )
            {
                HandleSmtpMail( j.ToString() );
            }//end of for				

            DateTime endTime = DateTime.Now;
            TimeSpan duration = endTime - startTime;
            string strTime = "Start time: " + startTime.ToString()
                + "\r\nEnd Time: " + endTime.ToString()
                + "\r\nDuration in Second: " + duration.TotalSeconds.ToString();
                            
            commObj.LogToFile( strTime );
            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
        }//end of Thd_SendSmtpMail

		/// <summary>
		/// Send Guid mail when user click on the send button
		/// Generate in threading manner for better user experience
		/// </summary>
		public void Thd_SendGuidMail()
		{
			Trace.WriteLine( "BatMailPage.cs - Thd_SendGuidMail" );
			
			this.Cursor = Cursors.WaitCursor;
			btnSend.Enabled = false;
            DateTime startTime = DateTime.Now;

			String inStr = txtSubject.Text; // save the user input
			System.Net.Sockets.TcpClient tcpClient = new TcpClient();
			try
			{		
                // open connection here - only open once
				tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboPort.Text) );			
				Debug.WriteLine("  +  Reading file - " + txtFrom.Text);
                // check number of repeat loop, and send until loop done
				for( int j = 0; j < nudLoop.Value; j++ )
				{
					HandleSendMail( inStr + j );
				}//end of for				
			}// end of try
			catch( System.Web.HttpException ex )
			{
				Trace.WriteLine( ex.Message.ToString() );
                commObj.LogToFile( "\tHttp Exception: " + ex.Message.ToString() );
			}//end of catch - generic exception
			finally
			{
				if( tcpClient != null )
				{
					Trace.WriteLine( "Finally - close TCP Clinet connection");
                    commObj.LogToFile( "Finally - close TCP Clinet connection" );
					tcpClient.Close();
				}//end of if 
			}//end of finally

            DateTime endTime = DateTime.Now;
            TimeSpan duration = endTime - startTime;
            string strTime = "Start time: " + startTime.ToString()
                + "\r\nEnd Time: " + endTime.ToString()
                + "\r\nDuration in Second: " + duration.TotalSeconds.ToString();
                            
            commObj.LogToFile( strTime );
			this.Cursor = Cursors.Default;
			btnSend.Enabled = true;		
		}//end of Thd_SendGuidMail

		/// <summary>
		/// Kill the send mail thread when program exit
		/// </summary>
		public void KillSendMailThread()
		{
			Trace.WriteLine("BatMailPage.cs - KillSendMailThread()");
			try
			{
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + sendMailThread.Name );
				sendMailThread.Abort(); // abort  
                sendMailThread.Join();  // require for ensure the thread kill

                // reset mouse cursor and enable send button ONLY Thread kill.
                this.Cursor = Cursors.Default;
                btnSend.Enabled = true;
			}//end of try 
			catch( ThreadAbortException thdEx )
			{
				Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in KillSendMailThread:" + sendMailThread.Name );
			}//end of catch				
		}//end of KillSendMailThread

        /// <summary>
        /// Kill the send mail thread when program exit
        /// </summary>
        public void KillFileMailThread()
        {
            Trace.WriteLine("BatMailPage.cs - KillFileMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + fileMailThread.Name );
                fileMailThread.Abort(); // abort  
                fileMailThread.Join();  // require for ensure the thread kill

                // reset mouse cursor and enable send button ONLY Thread kill.
                this.Cursor = Cursors.Default;
                btnSend.Enabled = true;
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in KillFileMailThread:" + fileMailThread.Name );
            }//end of catch				
        }//end of KillSendMailThread

        /// <summary>
        /// Kill the SMTP mail thread when program exit
        /// </summary>
        public void KillSmtpMailThread()
        {
            Trace.WriteLine("BatMailPage.cs - KillSmtpMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + smtpMailThread.Name );
                smtpMailThread.Abort(); // abort
                smtpMailThread.Join();  // require for ensure the thread kill

                // reset mouse cursor and enable send button ONLY Thread kill.
                this.Cursor = Cursors.Default;
                btnSend.Enabled = true;
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in KillSmtpMailThread:" + smtpMailThread.Name );
            }//end of catch				
        }//end of KillSmtpMailThread
        
        private void rdoNormal_Click(object sender, System.EventArgs e)
		{
			// enable normal group control
			rdoNormal.Checked = true;
			lnkFrom.Enabled = true;
			lnkTo.Enabled   = true;
			lnkCC.Enabled   = true;
			lnkBCC.Enabled  = true;
			txtFrom.Enabled = true;
			cboTo.Enabled   = true;
			cboCC.Enabled   = true;
			cboBCC.Enabled  = true;

            // disable SMTP group control
            rdoSMTP.Checked       = false;
            lblRcptTo.Enabled     = false;
            lblMailFrom.Enabled   = false;
            lnkFileStream.Enabled = false;
            cboRcptTo.Enabled     = false;
            cboMailFrom.Enabled   = false;
            txtInputFile.Enabled  = false;
            
            // disable outlook case control
			rdoOLCase.Checked = false;
			lnkFile.Enabled   = false;
			txtMailAddrFile.Enabled = false;

            // enable other control not used
            chkGUID.Enabled    = true;
            chkAttach.Enabled  = true;
            lblSubject.Enabled = true;
            lnkFolder.Enabled  = true;
            txtSubject.Enabled = true;
            txtFolder.Enabled  = true;
		}//end of rdoNormal_Click

		private void rdoOLCase_Click(object sender, System.EventArgs e)
		{
			// disable normal group control
			rdoNormal.Checked = false;
			lnkFrom.Enabled = false;
			lnkTo.Enabled   = false;
			lnkCC.Enabled   = false;
			lnkBCC.Enabled  = false;
			txtFrom.Enabled = false;
			cboTo.Enabled   = false;
			cboCC.Enabled   = false;
			cboBCC.Enabled  = false;

            // disable SMTP group control
            rdoSMTP.Checked       = false;
            lblRcptTo.Enabled     = false;
            lblMailFrom.Enabled   = false;
            lnkFileStream.Enabled = false;
            cboRcptTo.Enabled     = false;
            cboMailFrom.Enabled   = false;
            txtInputFile.Enabled  = false;

			// enable outlook case control
			rdoOLCase.Checked = true;
			lnkFile.Enabled   = true;
			txtMailAddrFile.Enabled = true;	

            // enable other control not used
            chkGUID.Enabled    = true;
            chkAttach.Enabled  = true;
            lblSubject.Enabled = true;
            lnkFolder.Enabled  = true;
            txtSubject.Enabled = true;
            txtFolder.Enabled  = true;

		}// end of rdoOLCase_Click

        private void rdoSMTP_Click(object sender, System.EventArgs e)
        {
            // disable normal group control
            rdoNormal.Checked = false;
            lnkFrom.Enabled = false;
            lnkTo.Enabled   = false;
            lnkCC.Enabled   = false;
            lnkBCC.Enabled  = false;
            txtFrom.Enabled = false;
            cboTo.Enabled   = false;
            cboCC.Enabled   = false;
            cboBCC.Enabled  = false;

            // disable SMTP group control
            rdoSMTP.Checked       = true;
            lblRcptTo.Enabled     = true;
            lblMailFrom.Enabled   = true;
            lnkFileStream.Enabled = true;
            cboRcptTo.Enabled     = true;
            cboMailFrom.Enabled   = true;
            txtInputFile.Enabled  = true;

            // disable outlook case control
            rdoOLCase.Checked = false;
            lnkFile.Enabled   = false;
            txtMailAddrFile.Enabled = false;

            // disable other control not used
            chkGUID.Enabled    = false;
            chkAttach.Enabled  = false;
            lblSubject.Enabled = false;
            lnkFolder.Enabled  = false;
            txtSubject.Enabled = false;
            txtFolder.Enabled  = false;
        
        }// end of rdoSMTP_Click

        public void Thd_SendFileMail()
        {
            Trace.WriteLine( "BatMailPage.cd - Thd_SendFileMail");
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;

            txtMailAddrFile.Text.TrimStart( new char[] {' '} );
            if( txtMailAddrFile.Text.Length <= 0 )
                return;

            // file name exist - do the job
            int    counter = 0; //mail sent counter
            string savSubj = txtSubject.Text; // save the user input subject
            string savBox  = richBox.Text;    // custom info in rich text box
            for( int j = 0; j < nudLoop.Value; j++ )
            {
                string strGUID = "";
                string strLine; // read from file
                StreamReader sr = null;	

                try
                {
                    sr = new StreamReader( txtMailAddrFile.Text ); // address book - put here for exception catch
                    while( (strLine = sr.ReadLine()) != null ) // file name from txtFrom field
                    {
                        counter++;
                        Debug.WriteLine( "\t - AutoSendMail - inside while loop : " + counter.ToString() );

                        if( strLine[0] != '#' ) // skip all comment
                        {
                            richBox.Text += "\r\nRead line : " + strLine;
                            if( chkGUID.Checked )
                            {
                                strGUID = System.Guid.NewGuid().ToString();
                                txtSubject.Text = savSubj + counter.ToString() + " " + strGUID;
                                commObj.LogGUID( "GUID.LOG", strGUID );
                            }//end of if - GUID			

                            // parse each line from the file, which defines a specific field:
                            // Total 4 fields: FROM, TO, CC, BCC and separated by commer.
                            // Fill can be null, and store in an array.
                            string [] splitStr = new string[4];						
                            splitStr = strLine.Split( new Char [] {','} );
                            for( int k = 0; k < splitStr.Length; k++ ) // trim leading and ending space
                                splitStr[k] = splitStr[k].Trim(' ');

                            MailMessage mailMsg = new MailMessage();

                            mailMsg.From	= splitStr[0];	// single line from file
                            mailMsg.To		= splitStr[1];	// single line from GUI
                            mailMsg.Cc		= splitStr[2];	// user input, may != To
                            mailMsg.Bcc		= splitStr[3];	// user input, may != To
                            mailMsg.Subject = txtSubject.Text;
                            mailMsg.Body	= savBox		// unchange user info
                                + "\n" + DateTime.Now
                                + "\n Mail Count == " + counter.ToString()
                                + "\nFrom: " + splitStr[0]	// start from here change
                                + "\n TO:  " + splitStr[1]
                                + "\n CC:  " + splitStr[2]
                                + "\n BCC: " + splitStr[3]
                                + "\n Body_Subject: " + txtSubject.Text;								

                            if( chkAttach.Checked )
                            {									
                                if( attachObj.idxAttach == attachObj.numFile )
                                    attachObj.idxAttach = 0; //reset

                                Debug.WriteLine( attachObj.idxAttach, "\t - idxAttach" );
                                Debug.WriteLine( attachObj.attchFullName, "\t - filename" );
                                mailMsg.Attachments.Add( new MailAttachment( attachObj.attchFullName, MailEncoding.Base64 ) );								

                                mailMsg.Body += "\r\nAttach file index = " + attachObj.idxAttach
                                    + "\r\nAttach file name = " + attachObj.attchFullName;

                                attachObj.idxAttach++; // point to next file
                            }//end of if

                            try
                            {
                                Debug.WriteLine("  +  Auto send mail la");
                                richBox.Text += "\r\n+ Do the send mail";
                                
                                // If SmtpServer is not set, local SMTP server is used
                                SmtpMail.SmtpServer = cboSMTP.Text; // port number??
                                SmtpMail.Send( mailMsg );
                                
                                richBox.Text = "Total Mail Sent: " + counter.ToString();
                            }//end of try
                            catch( System.Web.HttpException ex )
                            {
                                Debug.WriteLine(ex.Message.ToString());
                                commObj.LogGUID( "GUID.LOG", ex.Message.ToString() );
                                commObj.LogToFile( "error - " + ex.Message.ToString() );
                            }// end of catch

                            richBox.Text += "\r\n+ Batch Mails Sent Info " + mailMsg.Body; // display in rich Box
#if(DEBUG)
                            commObj.LogToFile( richBox.Text ); // save into log
#endif
                        }//end of if - skip all commtn

                        Thread.Sleep( (int)nudDelay.Value * 1000 );
                    }//end of while
                }//end of try
                catch( Exception ex )
                {
                    Debug.WriteLine(ex.Message.ToString());
                    MessageBox.Show(ex.Message.ToString(), msgCaption);
                    commObj.LogToFile( ex.Message.ToString() ); // save into log
                }//end of catch
                finally
                {
                    if( sr != null )
                    {
                        Trace.WriteLine("Finally - close the Stream Reader");
                        sr.Close();
                    }//end of if
                }//end of finally
            }//end of for

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;		
        }//end of Thd_SendFileMail

        #region Delete later
		/// <summary>
		/// Automatically send mail base on an input file.
		/// Input file format:
		/// Total 4 fields: FROM, TO, CC, BCC and separated by commer.
		/// </summary>
		/// 
/******
		public void AutoSendMail()
		{
			this.Cursor = Cursors.WaitCursor;
			btnSend.Enabled = false;

			txtMailAddrFile.Text.TrimStart( new char[] {' '} );
			if( txtMailAddrFile.Text.Length <= 0 )
				return;

			// file name exist - do the job
            int    counter = 0; //mail sent counter
			string savSubj = txtSubject.Text; // save the user input subject
			string savBox  = richBox.Text;    // custom info in rich text box
			for( int j = 0; j < nudLoop.Value; j++ )
			{
				string strGUID = "";
				string strLine; // read from file
				StreamReader sr = null;	

				try
				{
					sr = new StreamReader( txtMailAddrFile.Text ); // address book - put here for exception catch
					while( (strLine = sr.ReadLine()) != null ) // file name from txtFrom field
					{
                        counter++;
						Debug.WriteLine( "\t - AutoSendMail - inside while loop : " + counter.ToString() );

						if( strLine[0] != '#' ) // skip all comment
						{
							richBox.Text += "\r\nRead line : " + strLine;
							if( chkGUID.Checked )
							{
								strGUID = System.Guid.NewGuid().ToString();
								txtSubject.Text = savSubj + counter.ToString() + " " + strGUID;
								commObj.LogGUID( "GUID.LOG", strGUID );
							}//end of if - GUID			

							// parse each line from the file, which defines a specific field:
							// Total 4 fields: FROM, TO, CC, BCC and separated by commer.
							// Fill can be null, and store in an array.
							string [] splitStr = new string[4];						
							splitStr = strLine.Split( new Char [] {','} );
							for( int k = 0; k < splitStr.Length; k++ ) // trim leading and ending space
								splitStr[k] = splitStr[k].Trim(' ');

							MailMessage mailMsg = new MailMessage();

							mailMsg.From	= splitStr[0];	// single line from file
							mailMsg.To		= splitStr[1];	// single line from GUI
							mailMsg.Cc		= splitStr[2];	// user input, may != To
							mailMsg.Bcc		= splitStr[3];	// user input, may != To
							mailMsg.Subject = txtSubject.Text;
							mailMsg.Body	= savBox		// unchange user info
                                + "\n" + DateTime.Now
                                + "\n Mail Count == " + counter.ToString()
								+ "\nFrom: " + splitStr[0]	// start from here change
								+ "\n TO:  " + splitStr[1]
								+ "\n CC:  " + splitStr[2]
								+ "\n BCC: " + splitStr[3]
								+ "\n Body_Subject: " + txtSubject.Text;								

							if( chkAttach.Checked )
							{									
								if( attachObj.idxAttach == attachObj.numFile )
									attachObj.idxAttach = 0; //reset

								Debug.WriteLine( attachObj.idxAttach, "\t - idxAttach" );
								Debug.WriteLine( attachObj.attchFullName, "\t - filename" );
								mailMsg.Attachments.Add( new MailAttachment( attachObj.attchFullName, MailEncoding.Base64 ) );								

								mailMsg.Body += "\r\nAttach file index = " + attachObj.idxAttach
									+ "\r\nAttach file name = " + attachObj.attchFullName;

								attachObj.idxAttach++; // point to next file
							}//end of if

							try
							{
								Debug.WriteLine("  +  Auto send mail la");
								richBox.Text += "\r\n+ Do the send mail";
                                
                                // If SmtpServer is not set, local SMTP server is used
                                SmtpMail.SmtpServer = cboSMTP.Text; // port number??
								SmtpMail.Send( mailMsg );
                                
                                richBox.Text = "Total Mail Sent: " + counter.ToString();
							}//end of try
							catch( System.Web.HttpException ex )
							{
								Debug.WriteLine(ex.Message.ToString());
								commObj.LogGUID( "GUID.LOG", ex.Message.ToString() );
								commObj.LogToFile( "error - " + ex.Message.ToString() );
							}// end of catch

							richBox.Text += "\r\n+ Batch Mails Sent Info " + mailMsg.Body; // display in rich Box
#if(DEBUG)
							commObj.LogToFile( richBox.Text ); // save into log
#endif
						}//end of if - skip all commtn

						Thread.Sleep( (int)nudDelay.Value * 1000 );
					}//end of while
				}//end of try
				catch( Exception ex )
				{
					Debug.WriteLine(ex.Message.ToString());
					MessageBox.Show(ex.Message.ToString(), msgCaption);
                    commObj.LogToFile( ex.Message.ToString() ); // save into log
				}//end of catch
				finally
				{
					if( sr != null )
					{
						Trace.WriteLine("Finally - close the Stream Reader");
						sr.Close();
					}//end of if
				}//end of finally
			}//end of for

			this.Cursor = Cursors.Default;
			btnSend.Enabled = true;		
		}//end of AutoSendMail
****/
        #endregion

		private void lnkFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatMailPage.cs - lnkFile_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtMailAddrFile.Text = ofDlg.FileName;
			}//end of if				
		}//end of lnkFile_LinkClicked

        private void lnkFileStream_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "BatMailPage.cs - lnkFileStream_LinkClicked" );

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.ShowReadOnly = true;
            ofDlg.RestoreDirectory = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtInputFile.Text = ofDlg.FileName;
            }//end of if				        
        }// end of lnkFileStream_LinkClicked

        private void btnAbort_Click(object sender, System.EventArgs e)
        {
            try
            {
                if( sendMailThread != null && sendMailThread.IsAlive )
                    this.KillSendMailThread();

                if( smtpMailThread != null && smtpMailThread.IsAlive )
                    this.KillSmtpMailThread();  
      
                if( fileMailThread != null && fileMailThread.IsAlive )
                    this.KillFileMailThread();

            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine("BatMailPage.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                commObj.LogToFile("BatMailPage.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                MessageBox.Show( ex.Message + "\n" + ex.StackTrace, "Abort Exception" );
            }//end of catch
        }// end of lnkFileStream_LinkClicked
	} //end of class - BatMailPage
}//end of name space - QATool
