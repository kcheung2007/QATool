using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for OutlookPage.
	/// This is a new approach that separate the UI code and the outlook implementation code
	/// </summary>
	public class OutlookPage : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.ComboBox cboBCC;
		private System.Windows.Forms.ComboBox cboCC;
		private System.Windows.Forms.CheckBox chkGUID;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.Button btnSend;
		private System.Windows.Forms.CheckBox chkAttach;
		private System.Windows.Forms.NumericUpDown nudLoop;
		private System.Windows.Forms.Label lblLoop;
		private System.Windows.Forms.Label lblSubject;
		private System.Windows.Forms.TextBox txtSubject;
		private System.Windows.Forms.LinkLabel lnkBCC;
		private System.Windows.Forms.LinkLabel lnkCC;
		private System.Windows.Forms.RichTextBox richBox;
		private System.ComponentModel.IContainer components;

		private Thread olMailThread;
		private QATool.CommObj commObj = new CommObj();
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.LinkLabel lnkProfile;
        private System.Windows.Forms.ComboBox cboProfile;
		private QATool.AttachObj attachObj = null;
		private System.Windows.Forms.ComboBox cboTo;
		private System.Windows.Forms.LinkLabel lnkTo;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.RadioButton rdoFile;
		private System.Windows.Forms.RadioButton rdoUI;
		private System.Windows.Forms.ToolTip ttipOLPage;
		private System.Windows.Forms.TextBox txtFile;
		private System.Windows.Forms.LinkLabel lnkFolder;
		private System.Windows.Forms.CheckBox chkMultiAttach;
		private System.Windows.Forms.LinkLabel lnkAttach;
		private System.Windows.Forms.TextBox txtAttach;
		private System.Windows.Forms.Label lblDelay;
		private System.Windows.Forms.NumericUpDown nudDelay;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.LinkLabel lnkFile;

		public OutlookPage()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            commObj.InitComboBoxItem( cboProfile, "[Profile]" );
			commObj.InitComboBoxItem( cboTo, "[To Address]" );
			commObj.InitComboBoxItem( cboCC, "[CC Address]" );
			commObj.InitComboBoxItem( cboBCC,"[BCC Address]" );

		}// end of constructor

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				Debug.WriteLine( "outlookPage.cs - Deposing OutlookPage Object");
                if( olMailThread != null && olMailThread.IsAlive )
                {
                    this.KillolMailThread();
                    commObj.LogToFile( "Thread.log", "++ OutlookMailThread Killed ++");
                }

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
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.chkGUID = new System.Windows.Forms.CheckBox();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.chkAttach = new System.Windows.Forms.CheckBox();
            this.nudLoop = new System.Windows.Forms.NumericUpDown();
            this.lblLoop = new System.Windows.Forms.Label();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.lnkFile = new System.Windows.Forms.LinkLabel();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lnkProfile = new System.Windows.Forms.LinkLabel();
            this.cboProfile = new System.Windows.Forms.ComboBox();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoFile = new System.Windows.Forms.RadioButton();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdoUI = new System.Windows.Forms.RadioButton();
            this.txtAttach = new System.Windows.Forms.TextBox();
            this.chkMultiAttach = new System.Windows.Forms.CheckBox();
            this.lnkAttach = new System.Windows.Forms.LinkLabel();
            this.ttipOLPage = new System.Windows.Forms.ToolTip(this.components);
            this.nudDelay = new System.Windows.Forms.NumericUpDown();
            this.lblDelay = new System.Windows.Forms.Label();
            this.btnCheck = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).BeginInit();
            this.SuspendLayout();
            // 
            // cboBCC
            // 
            this.cboBCC.Location = new System.Drawing.Point(80, 156);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(296, 21);
            this.cboBCC.TabIndex = 66;
            this.ttipOLPage.SetToolTip(this.cboBCC, "bcc to");
            // 
            // cboCC
            // 
            this.cboCC.Location = new System.Drawing.Point(80, 132);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(296, 21);
            this.cboCC.TabIndex = 65;
            this.ttipOLPage.SetToolTip(this.cboCC, "cc to");
            // 
            // chkGUID
            // 
            this.chkGUID.Location = new System.Drawing.Point(12, 240);
            this.chkGUID.Name = "chkGUID";
            this.chkGUID.Size = new System.Drawing.Size(92, 16);
            this.chkGUID.TabIndex = 63;
            this.chkGUID.Text = "Include GUID";
            this.ttipOLPage.SetToolTip(this.chkGUID, "include GUID");
            // 
            // txtFolder
            // 
            this.txtFolder.Enabled = false;
            this.txtFolder.Location = new System.Drawing.Point(140, 36);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(232, 20);
            this.txtFolder.TabIndex = 62;
            this.txtFolder.Text = "C:\\C#Proj\\QATool";
            this.ttipOLPage.SetToolTip(this.txtFolder, "Attachment folder ONLY");
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(312, 236);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(64, 21);
            this.btnSend.TabIndex = 60;
            this.btnSend.Text = "Send";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // chkAttach
            // 
            this.chkAttach.Enabled = false;
            this.chkAttach.Location = new System.Drawing.Point(8, 40);
            this.chkAttach.Name = "chkAttach";
            this.chkAttach.Size = new System.Drawing.Size(80, 16);
            this.chkAttach.TabIndex = 59;
            this.chkAttach.Text = "Attachment";
            this.ttipOLPage.SetToolTip(this.chkAttach, "Include Attachment");
            this.chkAttach.CheckedChanged += new System.EventHandler(this.chkAttach_CheckedChanged);
            // 
            // nudLoop
            // 
            this.nudLoop.Location = new System.Drawing.Point(244, 236);
            this.nudLoop.Maximum = new System.Decimal(new int[] {
                                                                    9999,
                                                                    0,
                                                                    0,
                                                                    0});
            this.nudLoop.Minimum = new System.Decimal(new int[] {
                                                                    1,
                                                                    0,
                                                                    0,
                                                                    0});
            this.nudLoop.Name = "nudLoop";
            this.nudLoop.Size = new System.Drawing.Size(64, 20);
            this.nudLoop.TabIndex = 58;
            this.nudLoop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttipOLPage.SetToolTip(this.nudLoop, "0..9999");
            this.nudLoop.Value = new System.Decimal(new int[] {
                                                                  1,
                                                                  0,
                                                                  0,
                                                                  0});
            // 
            // lblLoop
            // 
            this.lblLoop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblLoop.Location = new System.Drawing.Point(208, 240);
            this.lblLoop.Name = "lblLoop";
            this.lblLoop.Size = new System.Drawing.Size(32, 16);
            this.lblLoop.TabIndex = 57;
            this.lblLoop.Text = "Loop";
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(20, 216);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 56;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(80, 212);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(296, 20);
            this.txtSubject.TabIndex = 55;
            this.txtSubject.Text = "txtSubject";
            // 
            // lnkBCC
            // 
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(40, 160);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(36, 20);
            this.lnkBCC.TabIndex = 53;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkCC
            // 
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(48, 136);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(28, 16);
            this.lnkCC.TabIndex = 52;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkFile
            // 
            this.lnkFile.Enabled = false;
            this.lnkFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFile.Location = new System.Drawing.Point(40, 14);
            this.lnkFile.Name = "lnkFile";
            this.lnkFile.Size = new System.Drawing.Size(32, 16);
            this.lnkFile.TabIndex = 51;
            this.lnkFile.TabStop = true;
            this.lnkFile.Text = "File :";
            this.lnkFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttipOLPage.SetToolTip(this.lnkFile, "Browse the address file");
            this.lnkFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFile_LinkClicked);
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(6, 264);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(302, 156);
            this.richBox.TabIndex = 67;
            this.richBox.Text = "richBox";
            // 
            // txtFile
            // 
            this.txtFile.Enabled = false;
            this.txtFile.Location = new System.Drawing.Point(76, 12);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(296, 20);
            this.txtFile.TabIndex = 54;
            this.txtFile.Text = "load address from file";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(228, 12);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '+';
            this.txtPassword.Size = new System.Drawing.Size(144, 20);
            this.txtPassword.TabIndex = 69;
            this.txtPassword.Text = "password0";
            this.ttipOLPage.SetToolTip(this.txtPassword, "password");
            // 
            // lnkProfile
            // 
            this.lnkProfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkProfile.Location = new System.Drawing.Point(28, 16);
            this.lnkProfile.Name = "lnkProfile";
            this.lnkProfile.Size = new System.Drawing.Size(44, 16);
            this.lnkProfile.TabIndex = 68;
            this.lnkProfile.TabStop = true;
            this.lnkProfile.Text = "Profile :";
            this.lnkProfile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboProfile
            // 
            this.cboProfile.Location = new System.Drawing.Point(76, 12);
            this.cboProfile.Name = "cboProfile";
            this.cboProfile.Size = new System.Drawing.Size(148, 21);
            this.cboProfile.TabIndex = 70;
            this.ttipOLPage.SetToolTip(this.cboProfile, "outlook profile");
            // 
            // cboTo
            // 
            this.cboTo.Items.AddRange(new object[] {
                                                       ""});
            this.cboTo.Location = new System.Drawing.Point(80, 108);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(296, 21);
            this.cboTo.TabIndex = 72;
            this.cboTo.Text = "login0@company1.zantaz.com";
            this.ttipOLPage.SetToolTip(this.cboTo, "mail to ");
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(48, 112);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(28, 16);
            this.lnkTo.TabIndex = 71;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoFile);
            this.groupBox1.Controls.Add(this.lnkFile);
            this.groupBox1.Controls.Add(this.txtFile);
            this.groupBox1.Controls.Add(this.chkAttach);
            this.groupBox1.Controls.Add(this.txtFolder);
            this.groupBox1.Controls.Add(this.lnkFolder);
            this.groupBox1.Location = new System.Drawing.Point(4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(380, 64);
            this.groupBox1.TabIndex = 73;
            this.groupBox1.TabStop = false;
            // 
            // rdoFile
            // 
            this.rdoFile.AutoCheck = false;
            this.rdoFile.Location = new System.Drawing.Point(8, 14);
            this.rdoFile.Name = "rdoFile";
            this.rdoFile.Size = new System.Drawing.Size(16, 16);
            this.rdoFile.TabIndex = 55;
            this.ttipOLPage.SetToolTip(this.rdoFile, "Automate from file");
            this.rdoFile.Click += new System.EventHandler(this.rdoFile_Click);
            // 
            // lnkFolder
            // 
            this.lnkFolder.Enabled = false;
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.Location = new System.Drawing.Point(96, 40);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(40, 16);
            this.lnkFolder.TabIndex = 61;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttipOLPage.SetToolTip(this.lnkFolder, "Locate the attachement folder");
            this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdoUI);
            this.groupBox2.Controls.Add(this.txtPassword);
            this.groupBox2.Controls.Add(this.lnkProfile);
            this.groupBox2.Controls.Add(this.cboProfile);
            this.groupBox2.Controls.Add(this.txtAttach);
            this.groupBox2.Controls.Add(this.chkMultiAttach);
            this.groupBox2.Controls.Add(this.lnkAttach);
            this.groupBox2.Location = new System.Drawing.Point(4, 68);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(380, 140);
            this.groupBox2.TabIndex = 74;
            this.groupBox2.TabStop = false;
            // 
            // rdoUI
            // 
            this.rdoUI.AutoCheck = false;
            this.rdoUI.Checked = true;
            this.rdoUI.Location = new System.Drawing.Point(8, 16);
            this.rdoUI.Name = "rdoUI";
            this.rdoUI.Size = new System.Drawing.Size(16, 16);
            this.rdoUI.TabIndex = 0;
            this.rdoUI.TabStop = true;
            this.ttipOLPage.SetToolTip(this.rdoUI, "Send individual mail");
            this.rdoUI.Click += new System.EventHandler(this.rdoUI_Click);
            // 
            // txtAttach
            // 
            this.txtAttach.Location = new System.Drawing.Point(76, 112);
            this.txtAttach.Name = "txtAttach";
            this.txtAttach.Size = new System.Drawing.Size(272, 20);
            this.txtAttach.TabIndex = 76;
            this.txtAttach.Text = "";
            this.ttipOLPage.SetToolTip(this.txtAttach, "Attachment - file name");
            // 
            // chkMultiAttach
            // 
            this.chkMultiAttach.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.chkMultiAttach.Location = new System.Drawing.Point(352, 116);
            this.chkMultiAttach.Name = "chkMultiAttach";
            this.chkMultiAttach.Size = new System.Drawing.Size(16, 16);
            this.chkMultiAttach.TabIndex = 77;
            this.chkMultiAttach.Text = "Multiple";
            this.chkMultiAttach.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.ttipOLPage.SetToolTip(this.chkMultiAttach, "Include multiple Attachment");
            // 
            // lnkAttach
            // 
            this.lnkAttach.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkAttach.Location = new System.Drawing.Point(8, 116);
            this.lnkAttach.Name = "lnkAttach";
            this.lnkAttach.Size = new System.Drawing.Size(68, 16);
            this.lnkAttach.TabIndex = 75;
            this.lnkAttach.TabStop = true;
            this.lnkAttach.Text = "Attachment : ";
            this.lnkAttach.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttipOLPage.SetToolTip(this.lnkAttach, "Browse attachment file");
            this.lnkAttach.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkAttach_LinkClicked);
            // 
            // nudDelay
            // 
            this.nudDelay.Location = new System.Drawing.Point(160, 236);
            this.nudDelay.Maximum = new System.Decimal(new int[] {
                                                                     60,
                                                                     0,
                                                                     0,
                                                                     0});
            this.nudDelay.Name = "nudDelay";
            this.nudDelay.Size = new System.Drawing.Size(44, 20);
            this.nudDelay.TabIndex = 76;
            this.nudDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttipOLPage.SetToolTip(this.nudDelay, "sec (0..5)");
            this.nudDelay.Value = new System.Decimal(new int[] {
                                                                   1,
                                                                   0,
                                                                   0,
                                                                   0});
            // 
            // lblDelay
            // 
            this.lblDelay.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblDelay.Location = new System.Drawing.Point(120, 240);
            this.lblDelay.Name = "lblDelay";
            this.lblDelay.Size = new System.Drawing.Size(36, 16);
            this.lblDelay.TabIndex = 75;
            this.lblDelay.Text = "Delay";
            this.ttipOLPage.SetToolTip(this.lblDelay, "in Sec (0..5)");
            // 
            // btnCheck
            // 
            this.btnCheck.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(192)), ((System.Byte)(128)));
            this.btnCheck.Font = new System.Drawing.Font("Arial Black", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.btnCheck.Location = new System.Drawing.Point(312, 396);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(68, 20);
            this.btnCheck.TabIndex = 77;
            this.btnCheck.Text = "ID Check";
            this.ttipOLPage.SetToolTip(this.btnCheck, "Launch checker window");
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // OutlookPage
            // 
            this.Controls.Add(this.btnCheck);
            this.Controls.Add(this.nudDelay);
            this.Controls.Add(this.lblDelay);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.lnkTo);
            this.Controls.Add(this.richBox);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.chkGUID);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.nudLoop);
            this.Controls.Add(this.lblLoop);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lnkBCC);
            this.Controls.Add(this.lnkCC);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Name = "OutlookPage";
            this.Size = new System.Drawing.Size(388, 428);
            this.ttipOLPage.SetToolTip(this, "Outlook Type");
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

		private void btnSend_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine("OutlookPage.cs - btnSend_Click");		

			if( chkAttach.Checked )
			{
				QATool.AttachObj attachObj = new AttachObj( txtFolder.Text );
			}// end of if - attachment check

			olMailThread = new Thread( new ThreadStart(this.Thd_SendOutlookMail) );
			olMailThread.Name = "OutlookMailThread";
			olMailThread.Start();

            commObj.LogToFile( "Thread.log", "++ OutlookMailThread Start ++");
		}//end of btnSend_Click

		/// <summary>
		/// Send mail by usint outlook in threading manner
		/// </summary>
		private void Thd_SendOutlookMail()
		{
			Trace.WriteLine( "OutlookPage.cs - Thd_SendOutlookMail()" );
			this.Cursor = Cursors.WaitCursor;
			btnSend.Enabled = false;

			if( rdoUI.Checked ) // selected - info from UI
			{
				HandleUISendMail();
			}//end of if - select info from UI
			else
				if( rdoFile.Checked )
				{
					HandleFileSendMail();
				}//end of if - select info from file
			btnSend.Enabled = true;
			this.Cursor = Cursors.Default;
		}// end of Thd_SendOutlookMail

		/// <summary>
		/// Get the file name in which contain a list of mail addresses
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lnkFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "OutlookPage.cs - lnkFrom_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtFile.Text = ofDlg.FileName;
			}//end of if				
		}//end of lnkFile_LinkClicked

/******* duplicated - delete if everything OK.
		/// <summary>
		/// Kill the send mail thread when program exit
		/// </summary>
		public void KillSendMailThread()
		{
			Trace.WriteLine("BatMailPage.cs - KillSendMailThread()");
			try
			{
                commObj.LogToFile("\t++ Kill Thread:" + olMailThread.Name );
				olMailThread.Abort(); // abort
                olMailThread.Join();  // require for ensure the thread kill
			}//end of try 
			catch( ThreadAbortException thdEx )
			{
				Trace.WriteLine( thdEx.Message );
				commObj.LogToFile("\t++ Kill Thread Exception:" + thdEx.StackTrace.ToString() );
			}//end of catch				
		}// end of KillSendMailThread
***** duplicated ***/

		private void rdoFile_Click(object sender, System.EventArgs e)
		{
			rdoFile.Checked = true;
			rdoUI.Checked   = false;

			lnkFile.Enabled   = true;
			txtFile.Enabled   = true;
			chkAttach.Enabled = true;
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
					
			cboTo.Enabled          = false;
			cboCC.Enabled          = false;
			cboBCC.Enabled         = false;
			cboProfile.Enabled     = false;
			txtPassword.Enabled    = false;
			lnkProfile.Enabled     = false;
			lnkTo.Enabled          = false;
			lnkCC.Enabled          = false;
			lnkBCC.Enabled         = false;
            lnkAttach.Enabled      = false;
            txtAttach.Enabled      = false;
            chkMultiAttach.Enabled = false;

		} // end of rdoFile_Click

		private void rdoUI_Click(object sender, System.EventArgs e)
		{
			rdoFile.Checked   = false;
			rdoUI.Checked     = true;

			lnkFile.Enabled   = false;
			txtFile.Enabled   = false;
			chkAttach.Enabled = false;
			lnkFolder.Enabled = false;
			txtFolder.Enabled = false;
						
			cboTo.Enabled          = true;
			cboCC.Enabled          = true;
			cboBCC.Enabled         = true;
			cboProfile.Enabled     = true;
			txtPassword.Enabled    = true;		
			lnkProfile.Enabled     = true;
			lnkTo.Enabled          = true;
			lnkCC.Enabled          = true;
			lnkBCC.Enabled         = true;
            lnkAttach.Enabled      = true;
            txtAttach.Enabled      = true;
            chkMultiAttach.Enabled = true;

		}//end of rdoUI_Click

		/// <summary>
		/// Send outlook mail based on user input. Send one mail at a time.
		/// </summary>
		public void HandleUISendMail()
		{	
			QATool.olMailObj olmailObj = new olMailObj();
			for( int j = 0; j < nudLoop.Value; j++ )
			{
				string strGUID = "";
				try
				{
					if( chkGUID.Checked )
					{
						strGUID = System.Guid.NewGuid().ToString();
						txtSubject.Text = txtSubject.Text + j.ToString() + " " + strGUID;
						commObj.LogGUID( "olGUID.LOG", strGUID );
					}//end of if - GUID			

					olmailObj.strPassword = txtPassword.Text;
					olmailObj.strProfile  = cboProfile.Text;						
					olmailObj.strTo       = cboTo.Text;
					olmailObj.strCC       = cboCC.Text;
					olmailObj.strBCC      = cboBCC.Text;
					olmailObj.strSubject  = txtSubject.Text;
					olmailObj.strBody     = richBox.Text;

					// validate the input - trim the space before check the length
					txtAttach.Text.TrimStart( new char[] {' '} );
					if( 0 < txtAttach.Text.Length )
					{
						commObj.LogToFile( "Attachment - " + txtAttach.Text );
						olmailObj.strAttachName = txtAttach.Text;
					}//end of if - attachment
				
					olmailObj.dumpToOutbox();
				}//end of try
				catch( Exception ex )
				{
					MessageBox.Show(ex.Message.ToString(), "QATool");
					commObj.LogGUID( "olGUID.LOG", "Exception occur " + strGUID.ToString() + "\n\t" + ex.Message.ToString() );
				}//end of catch - exception
			}//end of for
		}//end of HandleUISendMail

		/// <summary>
		/// Read from a file for all senders, profiles and password.
		/// Then create a mail and dump into mail account (profile) outbox.
		/// </summary>
		public void HandleFileSendMail()
		{		
			commObj.LogToFile("OutlookPage.cs +++ Enter HandleFileSendMail() +++ Send mail file");
			QATool.olMailObj olmailObj = new olMailObj();

			if( chkAttach.Checked )
				attachObj = new AttachObj( txtFolder.Text );

            int    counter = 0; // mail sent count
			string tmpSubj = txtSubject.Text; // save the user input subject
			string tmpBody = richBox.Text;	  // save the rich Box info

#if(DEBUG)
			commObj.LogToFile("\t Ready get into for loop");
#endif
			for( int j = 0; j < nudLoop.Value; j++ )
			{
				string strGUID = "";
				string strLine = "";
				StreamReader sr = null;

				try
				{
					sr = new StreamReader( txtFile.Text ); // address book - put here for exception catch
#if(DEBUG)
					commObj.LogToFile("\t Inside for loop - create stream reader");
					commObj.LogToFile("\t StreamReader = " + sr.ToString() );
#endif

					while( (strLine = sr.ReadLine()) != null ) // file name from txtFrom field
					{
                        counter++;
						Debug.WriteLine( "\t - HandleFileSendMail - inside while loop : " + counter.ToString()  );
						if( strLine[0] != '#' ) // skip all comment
						{							
							richBox.Text = "\r\n- Read line : " + strLine;
							if( chkGUID.Checked )
							{
								strGUID = System.Guid.NewGuid().ToString();
								txtSubject.Text = tmpSubj + " " + strGUID;
//								commObj.LogGUID( "GUID.LOG", strGUID );
							}//end of if - GUID			

							string [] splitStr = new string[7];
						
							splitStr = strLine.Split( new Char [] {','} );
							for( int k = 0; k < splitStr.Length; k++ ) // trim leading and ending space
								splitStr[k] = splitStr[k].Trim(' ');
							
							if( splitStr.Length == 5 ) // must exactly 5 cloumn
							{
								olmailObj.strTo       = splitStr[0];
								olmailObj.strCC       = splitStr[1];
								olmailObj.strBCC      = splitStr[2];
								olmailObj.strProfile  = splitStr[3];
								olmailObj.strPassword = splitStr[4];

								olmailObj.strSubject  = txtSubject.Text;
								olmailObj.strBody     = tmpBody	// unchange user info
									+ "\n TO:  " + splitStr[0]
									+ "\n CC:  " + splitStr[1]
									+ "\n BCC: " + splitStr[2]
									+ "\n Body_Subject: " + txtSubject.Text
									+ "\n" + DateTime.Now;

								if( chkAttach.Checked )
								{																	
									if( attachObj.idxAttach == attachObj.numFile )
										attachObj.idxAttach = 0; //reset

									Debug.WriteLine( attachObj.idxAttach, "\t - idxAttach" );
									Debug.WriteLine( attachObj.attchFullName, "\t - filename" );
									olmailObj.strAttachName = attachObj.attchFullName;																

									olmailObj.strBody += "\r\nAttach file index = " + attachObj.idxAttach
													   + "\r\nAttach file name = " + attachObj.attchFullName;

									attachObj.idxAttach++; // point to next file
								}//end of if		
					
								bool flag = olmailObj.dumpToOutbox();

								if( chkGUID.Checked && flag )
								{	// log the guid after dumpToOutbox
									commObj.LogGUID( "GUID.LOG", strGUID );
								}//end of if - GUID
							}//end of if - correct file 
						}//end of if - skip comment                        
						Thread.Sleep( (int)nudDelay.Value * 1000 ); // 2 sec
					}//end of while

                    olmailObj.nsLogoff();
				}//end of try
				catch( Exception ex )
				{
					Debug.WriteLine( ex.Message.ToString(), "\t Exception" );					
					commObj.LogGUID( "GUID.LOG", "Exception occur " + strGUID.ToString() + "\n\t" + ex.Message.ToString() );
				}//end of catch - exception
				finally
				{
					if( sr != null )
					{
						Trace.WriteLine("Finally - close the Stream Reader");
						sr.Close();
					}//end of if
				}//end of finally - clean up everything
			}//end of for

			commObj.LogToFile("OutlookPage.cs +++ End of HandleFileSendMail() +++");
		}// end of HandleFileSendMail

		/// <summary>
		/// Kill the send mail thread when program exit
		/// </summary>
		public void KillolMailThread()
		{
			Trace.WriteLine("OutlookPage.cs - KillolMailThread()");
			try
			{
                commObj.LogToFile( "Thread.log", "Kill OutlookMailThread Start");
				olMailThread.Abort(); // abort
                olMailThread.Join();  // require for ensure the thread kill
			}//end of try 
			catch( ThreadAbortException thdEx )
			{
				Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Aborting the send guid mail thread : " + thdEx.Message.ToString() );

//				richBox.Text += "\nAborting the send guid mail thread";
			}//end of catch				
		}//end of KillolMailThread

		/// <summary>
		/// For the send mail from file section - not the GUI section
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
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
		} // end of chkAttach_CheckedChanged

		/// <summary>
		/// For send mail from file, browse the attachment folder in which contains files for attachment
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "OutlookPage.cs - lnkFolder_LinkClicked" );
			FolderBrowserDialog fbDlg = new FolderBrowserDialog();

			fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
			if( txtFolder.Text != null )
				fbDlg.SelectedPath = txtFolder.Text;  // set the default folder

			if( fbDlg.ShowDialog() == DialogResult.OK )
			{
				txtFolder.Text = fbDlg.SelectedPath;
			}//end of if
		}//end of lnkFolder_LinkClick

		/// <summary>
		/// 1) For GUI send mail, browse and select particular file.
		/// 2) Can attach multiple file base on the "multi" check box. 
		/// 3) Use semicolon for seperate multi files.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lnkAttach_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			string [] fileNames;

			OpenFileDialog ofDlg = new OpenFileDialog();
			if( chkMultiAttach.Checked )
			{
				ofDlg.Multiselect = true;
				if( ofDlg.ShowDialog() == DialogResult.OK )
				{
					fileNames = ofDlg.FileNames;
					foreach( string str in fileNames )
					{
						txtAttach.Text += ";" + str;
					}//end of foreach

					//check the first char
					string tmpStr = txtAttach.Text.ToString();
					if( tmpStr[0] == ';' )
						txtAttach.Text = txtAttach.Text.Remove(0,1);
				}//end of if		
			}//end of if
			else
			{
				if( ofDlg.ShowDialog() == DialogResult.OK )
				{
					txtAttach.Text = ofDlg.FileName;
				}//end of if		
			}//end of else		
		}//end of lnkAttach_LinkClicked

        /// <summary>
        /// Launch checker window to validate source GUID and result GUID
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheck_Click(object sender, System.EventArgs e)
        {
            CheckerWnd wndChecker = new CheckerWnd();
            wndChecker.ShowDialog();
        }//end of btnCheck_Click
	}
}
