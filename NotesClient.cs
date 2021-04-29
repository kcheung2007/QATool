using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for NotesClient.
    /// Register Lotus Notes object
    /// Open a command prompt window, change to the Notes program directory 
    /// and run "regsvr32 nlsxbe.dll". This registers the backend class library. 
    /// Be sure to set a reference to "Lotus Domino Objects" (domobj.tlb) to get fast vtable early binding support
	/// </summary>
	public class NotesClient : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.NumericUpDown nudDelay;
        private System.Windows.Forms.Label lblDelay;
        private System.Windows.Forms.ComboBox cboTo;
        private System.Windows.Forms.LinkLabel lnkTo;
        private System.Windows.Forms.ComboBox cboBCC;
        private System.Windows.Forms.RichTextBox richBox;
        private System.Windows.Forms.ComboBox cboCC;
        private System.Windows.Forms.CheckBox chkGUID;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.LinkLabel lnkCC;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoFile;
        private System.Windows.Forms.LinkLabel lnkFile;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.CheckBox chkAttach;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.LinkLabel lnkFolder;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdoUI;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtAttach;
        private System.Windows.Forms.CheckBox chkMultiAttach;
        private System.Windows.Forms.LinkLabel lnkAttach;
        private System.Windows.Forms.NumericUpDown nudLoop;
        private System.Windows.Forms.Label lblLoop;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.LinkLabel lnkBCC;
        private System.Windows.Forms.ToolTip ttipOLPage;
        private System.Windows.Forms.ComboBox cboNotesItem;
        private System.Windows.Forms.Label lblNotesItem;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.Button btnAbort;

        private QATool.CommObj commObj = new CommObj();        
        private Thread ncMailThread;
        private QATool.AttachObj attachObj = null;

		public NotesClient()
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
                Debug.WriteLine( "NotesClient.cs - Deposing NotesClient Page Object");
                if( ncMailThread != null && ncMailThread.IsAlive )
                {
                    this.KillNotesClientMailThread();
                    commObj.LogToFile( "Thread.log", "++ KillNotesClientMailThread Killed ++");
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
            this.nudDelay = new System.Windows.Forms.NumericUpDown();
            this.lblDelay = new System.Windows.Forms.Label();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.chkGUID = new System.Windows.Forms.CheckBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoFile = new System.Windows.Forms.RadioButton();
            this.lnkFile = new System.Windows.Forms.LinkLabel();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.chkAttach = new System.Windows.Forms.CheckBox();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblNotesItem = new System.Windows.Forms.Label();
            this.rdoUI = new System.Windows.Forms.RadioButton();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.cboNotesItem = new System.Windows.Forms.ComboBox();
            this.txtAttach = new System.Windows.Forms.TextBox();
            this.chkMultiAttach = new System.Windows.Forms.CheckBox();
            this.lnkAttach = new System.Windows.Forms.LinkLabel();
            this.nudLoop = new System.Windows.Forms.NumericUpDown();
            this.lblLoop = new System.Windows.Forms.Label();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.ttipOLPage = new System.Windows.Forms.ToolTip(this.components);
            this.btnAbort = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).BeginInit();
            this.SuspendLayout();
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
            this.nudDelay.TabIndex = 94;
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
            this.lblDelay.TabIndex = 93;
            this.lblDelay.Text = "Delay";
            this.ttipOLPage.SetToolTip(this.lblDelay, "in Sec (0..5)");
            // 
            // cboTo
            // 
            this.cboTo.Items.AddRange(new object[] {
                                                       ""});
            this.cboTo.Location = new System.Drawing.Point(80, 108);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(296, 21);
            this.cboTo.TabIndex = 90;
            this.cboTo.Text = "admin@zel.zantaz.com";
            this.ttipOLPage.SetToolTip(this.cboTo, "mail to (separated by comma)");
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(48, 112);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(28, 16);
            this.lnkTo.TabIndex = 89;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboBCC
            // 
            this.cboBCC.Location = new System.Drawing.Point(80, 156);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(296, 21);
            this.cboBCC.TabIndex = 87;
            this.ttipOLPage.SetToolTip(this.cboBCC, "separated by comma");
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(6, 264);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(302, 156);
            this.richBox.TabIndex = 88;
            this.richBox.Text = "richBox";
            // 
            // cboCC
            // 
            this.cboCC.Location = new System.Drawing.Point(80, 132);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(296, 21);
            this.cboCC.TabIndex = 86;
            this.ttipOLPage.SetToolTip(this.cboCC, "separated by comma");
            // 
            // chkGUID
            // 
            this.chkGUID.Location = new System.Drawing.Point(12, 240);
            this.chkGUID.Name = "chkGUID";
            this.chkGUID.Size = new System.Drawing.Size(92, 16);
            this.chkGUID.TabIndex = 85;
            this.chkGUID.Text = "Include GUID";
            this.ttipOLPage.SetToolTip(this.chkGUID, "include GUID");
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(312, 236);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(64, 21);
            this.btnSend.TabIndex = 84;
            this.btnSend.Text = "Send";
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // lnkCC
            // 
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(48, 136);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(28, 16);
            this.lnkCC.TabIndex = 78;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
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
            this.groupBox1.TabIndex = 91;
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
            // txtFile
            // 
            this.txtFile.Enabled = false;
            this.txtFile.Location = new System.Drawing.Point(76, 12);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(296, 20);
            this.txtFile.TabIndex = 54;
            this.txtFile.Text = "load address from file";
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
            this.groupBox2.Controls.Add(this.lblNotesItem);
            this.groupBox2.Controls.Add(this.rdoUI);
            this.groupBox2.Controls.Add(this.txtPassword);
            this.groupBox2.Controls.Add(this.cboNotesItem);
            this.groupBox2.Controls.Add(this.txtAttach);
            this.groupBox2.Controls.Add(this.chkMultiAttach);
            this.groupBox2.Controls.Add(this.lnkAttach);
            this.groupBox2.Location = new System.Drawing.Point(4, 68);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(380, 140);
            this.groupBox2.TabIndex = 92;
            this.groupBox2.TabStop = false;
            // 
            // lblNotesItem
            // 
            this.lblNotesItem.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblNotesItem.Location = new System.Drawing.Point(32, 16);
            this.lblNotesItem.Name = "lblNotesItem";
            this.lblNotesItem.Size = new System.Drawing.Size(40, 16);
            this.lblNotesItem.TabIndex = 96;
            this.lblNotesItem.Text = "Items";
            this.lblNotesItem.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttipOLPage.SetToolTip(this.lblNotesItem, "Notes Item");
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
            // cboNotesItem
            // 
            this.cboNotesItem.Location = new System.Drawing.Point(76, 12);
            this.cboNotesItem.Name = "cboNotesItem";
            this.cboNotesItem.Size = new System.Drawing.Size(148, 21);
            this.cboNotesItem.TabIndex = 70;
            this.cboNotesItem.Text = "memo";
            this.ttipOLPage.SetToolTip(this.cboNotesItem, "Notes Form type");
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
            this.nudLoop.TabIndex = 83;
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
            this.lblLoop.TabIndex = 82;
            this.lblLoop.Text = "Loop";
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(20, 216);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 81;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(80, 212);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(296, 20);
            this.txtSubject.TabIndex = 80;
            this.txtSubject.Text = "txtSubject";
            // 
            // lnkBCC
            // 
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(40, 160);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(36, 20);
            this.lnkBCC.TabIndex = 79;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnAbort
            // 
            this.btnAbort.Location = new System.Drawing.Point(312, 264);
            this.btnAbort.Name = "btnAbort";
            this.btnAbort.Size = new System.Drawing.Size(64, 21);
            this.btnAbort.TabIndex = 95;
            this.btnAbort.Text = "Abort";
            this.btnAbort.Click += new System.EventHandler(this.btnAbort_Click);
            // 
            // NotesClient
            // 
            this.Controls.Add(this.lnkBCC);
            this.Controls.Add(this.btnAbort);
            this.Controls.Add(this.nudDelay);
            this.Controls.Add(this.lblDelay);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.lnkTo);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.richBox);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.chkGUID);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.lnkCC);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.nudLoop);
            this.Controls.Add(this.lblLoop);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.txtSubject);
            this.Name = "NotesClient";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.nudDelay)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudLoop)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion


        /// <summary>
        /// Send mail by usint notes client in threading manner
        /// </summary>
        private void Thd_SendNotesClientMail()
        {
            Trace.WriteLine( "NotesClient.cs - Thd_SendNotesClientMail()" );
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
        /// Read from a file for all senders, receivers
        /// Text file format: To, CC, BCC, From, password.
        /// Each field separated by comma, mail addresses in each field separated by semi-colon;
        /// Two Loops:
        /// 1) Inner loop: sending mail based on the number of 'TO' users in the text file within a session.
        /// 2) Outer loop: repeat the whole mailing list in the file. Each loop will create a new session.
        /// Therefore, total number of sent mail == inner loop x outer loop.
        /// </summary>
        /// 11-29-04: From and password do not do anything
        public void HandleFileSendMail()
        {		
            Trace.WriteLine( "NotesClient.cs - HandleFileSendMail()" );

            int    counter = 0; // mail sent count
            string tmpSubj = txtSubject.Text; // save the user input subject
            string tmpBody = richBox.Text;	  // save the rich Box info

            string strTo  = "";
            string strCC  = "";
            string strBCC = "";

            string [] toArray;
            string [] ccArray;
            string [] bccArray;
            string inSubject  = "";
            string bodyText   = "";
            string notesItems = "";

            Domino.NotesDatabase     notesDB;
            Domino.NotesDocument     notesDoc;
            Domino.NotesItem         docForm;
            Domino.NotesItem         docSubject;
            Domino.NotesItem         docCopyTo;  // CC
            Domino.NotesItem         docBlindCC; // BCC
            Domino.NotesRichTextItem docRTFBody;                

            Object recipients;
            Object carboncopy;
            Object blindcopy;

            StreamReader sr = null;

//            StreamReader sr = new StreamReader( txtFile.Text ); // address book - put here for exception catch
            if( chkAttach.Checked ) // point to attachment folder ONLY
                attachObj = new AttachObj( txtFolder.Text );

            for( int j = 0; j < nudLoop.Value; j++ )
            {
                string strGUID = "";
                string strLine = "";                

                Trace.WriteLine("\t inside for loop: Open Domino Session");
                Domino.NotesSession domSession = new Domino.NotesSession();
                try
                {   // only with computer with Domino Server installed
                    // domSession.InitializeUsingNotesUserName("atest0", "password0");

                    // used on a computer with a Notes client/Domino server 
                    // and bases on the session on the current user ID - admin.id
                    // domSession.Initialize("");

                    // TO DO: Need to modify, should NOT reuse the UI password field.
                    domSession.Initialize( txtPassword.Text ); 
                }//end of try
                catch( Exception exSession )
                {
                    string msg = "Fail in creating Notes Session\n" + exSession.Message + "\n"
                        + exSession.GetType().ToString() + "\n" + exSession.StackTrace;
                    MessageBox.Show( msg, "QATool" );

                    Trace.WriteLine("\t Session Fail in HandelFileSendMail: releaseing session COM Object");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(domSession);
                }//end of catch - notes session exception

                try
                {
                    Domino.NotesDbDirectory domDBDir = domSession.GetDbDirectory("");
                    sr = new StreamReader( txtFile.Text ); // address book - put here for exception catch

                    Trace.WriteLine( "\t Inside for loop - create stream reader" );
                    Trace.WriteLine( "\t StreamReader = " + sr.ToString() );

                    while( (strLine = sr.ReadLine()) != null ) // file name from txtFrom field
                    {
                        counter++;
                        Trace.WriteLine( "\t - HandleFileSendMail - inside while loop : " + counter.ToString()  );
                        if( strLine[0] != '#' ) // skip all comment - first character
                        {							
                            richBox.Text = "\r\n- Read line : " + strLine;
                            if( chkGUID.Checked )
                            {
                                strGUID = System.Guid.NewGuid().ToString();
                                txtSubject.Text = tmpSubj + " " + strGUID;
 //								commObj.LogGUID( "GUID.LOG", strGUID );
                            }//end of if - GUID			

                            string [] splitStr = new string[5]; // match with the colum of input text file						
                            splitStr = strLine.Split( new Char [] {','} ); // each field separated by comma

                            Trace.WriteLine( "string[0] = " + splitStr[0].ToString() );
                            Trace.WriteLine( "string[1] = " + splitStr[1].ToString() );

                            for( int k = 0; k < splitStr.Length; k++ ) // trim leading and ending space
                            {
                                Trace.WriteLine( "k = " + k.ToString() + " splitStr = " + splitStr[k].ToString() );
                                splitStr[k] = splitStr[k].Trim(' ');
                            }//end of for
							
                            Trace.WriteLine( "after for loop - triming leading and ending space");
                            if( splitStr.Length == 5 ) // must exactly 5 cloumn
                            {
                                Trace.WriteLine( "split string == 5. Inside if");
                                strTo  = splitStr[0];
                                strCC  = splitStr[1];
                                strBCC = splitStr[2];

                                toArray  = strTo.Split(  new char [] { ';' } );
                                ccArray  = strCC.Split(  new char [] { ';' } );
                                bccArray = strBCC.Split( new char [] { ';' } );
                                inSubject  = counter.ToString() + " " + txtSubject.Text;
                                bodyText   = richBox.Text;
                                notesItems = "memo";

                                recipients = toArray;
                                carboncopy = ccArray;
                                blindcopy  = bccArray;

                                Trace.WriteLine( "Will open domino DB " );
                                notesDB    = domDBDir.OpenMailDatabase();
                                Trace.WriteLine( "After open domino DB " + notesDB.ToString());

                                notesDoc   = notesDB.CreateDocument();
                                docForm    = notesDoc.ReplaceItemValue("Form", notesItems);
                                docSubject = notesDoc.ReplaceItemValue("Subject", inSubject );
                                docCopyTo  = notesDoc.ReplaceItemValue("CopyTo", carboncopy);
                                docBlindCC = notesDoc.ReplaceItemValue("BlindCopyTo", blindcopy);
                                docRTFBody = notesDoc.CreateRichTextItem("Body");

                                bodyText   = richBox.Text 
                                    + "\n To: "  + cboTo.Text
                                    + "\n CC: "  + cboCC.Text
                                    + "\n BCC: " + cboBCC.Text
                                    + "\n Notes_Subject: " + inSubject
                                    + "\n " + DateTime.Now + "\n";
                                docRTFBody.AppendText( bodyText );

                                if( chkAttach.Checked )
                                {																	
                                    if( attachObj.idxAttach == attachObj.numFile )
                                        attachObj.idxAttach = 0; //reset

                                    Debug.WriteLine( attachObj.idxAttach, "\t - idxAttach" );
                                    Debug.WriteLine( attachObj.attchFullName, "\t - filename" );
                                    docRTFBody.EmbedObject( Domino.EMBED_TYPE.EMBED_ATTACHMENT,"", attachObj.attchFullName,"");

                                    bodyText += "\r\nAttach file index = " + attachObj.idxAttach
                                        + "\r\nAttach file name = " + attachObj.attchFullName;
                                }//end of if		
					
                                // update UI info:
                                txtSubject.Text = inSubject;
                                richBox.Text    = bodyText;

                                Trace.WriteLine( "Just before sending mail" );
                                notesDoc.Send(false, ref recipients); // send notes mail
                                Trace.WriteLine( "Just after sending mail" );

                                richBox.Text = ""; // clear UI richard box
                                txtSubject.Text = tmpSubj;
                                if( chkGUID.Checked )
                                {	// log the guid after dumpToOutbox
                                    commObj.LogGUID( "GUID.LOG", strGUID );
                                }//end of if - GUID
                            }//end of if - correct file 
                        }//end of if - skip comment                        
//                        Thread.Sleep( (int)nudDelay.Value * 1000 ); // 2 sec
                    }//end of while
                }//end of try
                catch( Exception ex )
                {
                    Trace.WriteLine( ex.Message.ToString(), "\t Exception" );					
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

            richBox.Text = "Total Mail send: " + counter.ToString();

            commObj.LogToFile("OutlookPage.cs +++ End of HandleFileSendMail() +++");
        }// end of HandleFileSendMail

        /// <summary>
        /// Send notes client mail based on user input from UI. Send one mail at a time.
        /// Only create one session and then loop through it.
        /// </summary>
        public void HandleUISendMail()
        {
            Debug.WriteLine("NotesClient.cs - HandleUISendMail");
            Domino.NotesSession domSession = null;
            int counter = 0;
            try
            {  
                domSession = new Domino.NotesSession();

                // only with computer with Domino Server installed
                // domSession.InitializeUsingNotesUserName("atest0", "password0");

                // used on a computer with a Notes client/Domino server 
                // and bases on the session on the current user ID - admin.id
                // domSession.Initialize("");
                domSession.Initialize( txtPassword.Text );
            }//end of try
            catch( Exception exSession )
            {
                string msg = "Fail in creating Notes Session\n" + exSession.Message + "\n"
                    + exSession.GetType().ToString() + "\n";
                MessageBox.Show( msg, "QATool" );
                Debug.WriteLine( msg + "\n"  + exSession.StackTrace );

                Debug.WriteLine("NotesClient.cs -   Session Fail: releaseing session COM Object");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(domSession);

                this.Cursor = Cursors.Default;  // reset cursor 
                btnSend.Enabled = true;         // re-enable send button
            }//end of catch - notes session exception

            string strGUID = "";
            string strTmpSubj = txtSubject.Text;
            try
            {
                Domino.NotesDbDirectory domDBDir = domSession.GetDbDirectory("");
                for( int j = 0; j < nudLoop.Value; j++ )
                {
                    if( chkGUID.Checked )
                    {
                        strGUID = System.Guid.NewGuid().ToString();
                        txtSubject.Text = txtSubject.Text + j.ToString() + " " + strGUID;
                        commObj.LogGUID( "olGUID.LOG", strGUID );
                    }//end of if - GUID

                    richBox.Text += "\nUser name: " + domSession.UserName
                        + "\nServer name: " + domSession.ServerName; // server name is null                

                    string strTo  = cboTo.Text;
                    string strCC  = cboCC.Text;
                    string strBCC = cboBCC.Text;

                    string [] toArray  = strTo.Split(  new char [] { ',', ';' } );
                    string [] ccArray  = strCC.Split(  new char [] { ',', ';' } );
                    string [] bccArray = strBCC.Split( new char [] { ',', ';' } );
                    string inSubject   = j.ToString() + " " + strTmpSubj;
                    string bodyText    = "";
                    string notesItems  = cboNotesItem.Text;

                    Domino.NotesDatabase     notesDB;
                    Domino.NotesDocument     notesDoc;
                    Domino.NotesItem         docForm;
                    Domino.NotesItem         docSubject;
                    Domino.NotesItem         docCopyTo;  // CC
                    Domino.NotesItem         docBlindCC; // BCC
                    Domino.NotesRichTextItem docRTFBody;                

                    Object recipients = toArray;
                    Object carboncopy = ccArray;
                    Object blindcopy  = bccArray;

                    notesDB    = domDBDir.OpenMailDatabase();
                    notesDoc   = notesDB.CreateDocument();

                    docForm    = notesDoc.ReplaceItemValue("Form", notesItems);
                    docSubject = notesDoc.ReplaceItemValue("Subject", inSubject );
                    docCopyTo  = notesDoc.ReplaceItemValue("CopyTo", carboncopy);
                    docBlindCC = notesDoc.ReplaceItemValue("BlindCopyTo", blindcopy);
                    docRTFBody = notesDoc.CreateRichTextItem("Body");

                    bodyText   = richBox.Text 
                        + "\n To: "  + cboTo.Text
                        + "\n CC: "  + cboCC.Text
                        + "\n BCC: " + cboBCC.Text
                        + "\n Notes_Subject: " + inSubject
                        + "\n " + DateTime.Now + "\n";



                    Trace.WriteLine( "Before docRTFBody.AppendText " + docRTFBody.Values.ToString() );
                    docRTFBody.AppendText( bodyText );
                    Trace.WriteLine( "docRTFBody.AppendText( bodyText )" );

                    string fn = txtAttach.Text;
                    if( fn != "" )
                    {
                        bodyText += fn;
                        char[] delim = new char[]{';'};
                        foreach( string str in fn.Split(delim) )
                        {
                            docRTFBody.EmbedObject( Domino.EMBED_TYPE.EMBED_ATTACHMENT,"",str,"");
                        }//end of foreach
                    }//end of if - adding attachment

                    // update UI info:
                    txtSubject.Text = inSubject;
                    richBox.Text    = bodyText;

                    counter++;
                    notesDoc.Send(false, ref recipients); // send notes mail
                    commObj.LogToFile( "Notes Document sent" );
                    richBox.Text = ""; // clear UI richard box
                }//end of for
                richBox.Text = "Total mail send = " + counter.ToString();
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message + "\n" + ex.GetType().ToString() + "\n" + ex.StackTrace;
                MessageBox.Show( msg, "Handle UI Info" );
            }//end of catch - exception
            finally
            {
                Debug.WriteLine("NotesClient.cs - finally - releaseing session COM Object");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(domSession);
            }//end of finally            
        }//end of HandleUISendMail

        private void btnSend_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine("NotesClient.cs - btnSend_Click");		
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;

            if( chkAttach.Checked )
            {
                QATool.AttachObj attachObj = new AttachObj( txtFolder.Text );
            }// end of if - attachment check

            ncMailThread = new Thread( new ThreadStart(this.Thd_SendNotesClientMail) );
            ncMailThread.Name = "NotesClientMailThread";
            ncMailThread.Start();

            commObj.LogToFile( "Thread.log", "++ NotesClientMailThread Start ++");
            
            btnSend.Enabled = true;
            this.Cursor = Cursors.Default;
        }// end of btnSend_Click

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
        /// Kill the send mail thread when program exit
        /// </summary>
        public void KillNotesClientMailThread()
        {
            Trace.WriteLine("NotesClient.cs - KillNotesClientMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "Kill Notes Client Mail Start");
                ncMailThread.Abort(); // abort
                ncMailThread.Join();  // require for ensure the thread kill
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Aborting the Notes Client Mail thread : " + thdEx.Message.ToString() );
            }//end of catch

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
        }//end of KillNotesClientMailThread

        private void btnAbort_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "NotesClient.cs - btnAbort_Click" );
            try
            {
                if( ncMailThread != null && ncMailThread.IsAlive )
                    KillNotesClientMailThread();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine("NotesClient.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                commObj.LogToFile("NotesClient.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                MessageBox.Show( ex.Message + "\n" + ex.StackTrace, "Abort Exception" );
            }//end of catch                
        }//end of btnAbort_Click

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
            txtPassword.Enabled    = true;		
            lnkTo.Enabled          = true;
            lnkCC.Enabled          = true;
            lnkBCC.Enabled         = true;    
            lblNotesItem.Enabled   = true;
            cboNotesItem.Enabled   = true;
            lnkAttach.Enabled      = true;
            txtAttach.Enabled      = true;
            chkMultiAttach.Enabled = true;
        }//end of rdoUI_Click

        private void rdoFile_Click(object sender, System.EventArgs e)
        {
            rdoFile.Checked   = true;
            rdoUI.Checked     = false;

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
            txtPassword.Enabled    = false;
            lnkTo.Enabled          = false;
            lnkCC.Enabled          = false;
            lnkBCC.Enabled         = false;
            lblNotesItem.Enabled   = false;
            cboNotesItem.Enabled   = false;
            lnkAttach.Enabled      = false;
            txtAttach.Enabled      = false;
            chkMultiAttach.Enabled = false;
        }//end of rdoFile_Click

        private void lnkFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "NotesClient.cs - lnkFrom_LinkClicked" );

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.ShowReadOnly = true;
            ofDlg.RestoreDirectory = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtFile.Text = ofDlg.FileName;
            }//end of if
        }//end of lnkFile_LinkClicked

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
        }// end of lnkFolder_LinkClicked

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
        }

	}
}
