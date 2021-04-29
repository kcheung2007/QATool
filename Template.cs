using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for Template.
	/// </summary>
	public class Template : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Label lblRcvdEnd;
		private System.Windows.Forms.Label lblNumTemplate;
		private System.Windows.Forms.TextBox txtPrefix;
		private System.Windows.Forms.NumericUpDown nudCount;
		private System.Windows.Forms.DateTimePicker dtpSendStart;
		private System.Windows.Forms.TextBox txtRepositID;
		private System.Windows.Forms.DateTimePicker dtpSendEnd;
		private System.Windows.Forms.ComboBox cboDocType;
		private System.Windows.Forms.GroupBox gbxCriteria1;
		private System.Windows.Forms.TextBox txtQ2;
		private System.Windows.Forms.TextBox txtTag2;
		private System.Windows.Forms.TextBox txtTag1;
		private System.Windows.Forms.TextBox txtQ1;
		private System.Windows.Forms.DateTimePicker dtpRcvdStart;
		private System.Windows.Forms.DateTimePicker dtpRcvdEnd;
		private System.Windows.Forms.Label lblRcvdStart;
		private System.Windows.Forms.ToolTip ttpAudit;
		private System.Windows.Forms.Button btnCreate;
		private System.Windows.Forms.GroupBox gbxCriteria2;
		private System.Windows.Forms.TextBox txtQ4;
		private System.Windows.Forms.TextBox txtQ3;
		private System.Windows.Forms.TextBox txtTag4;
		private System.Windows.Forms.TextBox txtTag3;
		private System.Windows.Forms.Label lblRepositID;
		private System.Windows.Forms.Label lblStartDate;
		private System.Windows.Forms.Label lblEndDate;
		private System.Windows.Forms.Label lblBox11;
		private System.Windows.Forms.Label lblBox12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label lblBox14;
		private System.Windows.Forms.Label lblBox21;
		private System.Windows.Forms.Label lblBox23;
		private System.Windows.Forms.Label lblBox22;
		private System.Windows.Forms.Label lblBox24;
		private System.Windows.Forms.Label lblPrefix;
		private System.Windows.Forms.Button btnModify;
		private System.Windows.Forms.Button btnFolder;
		private System.ComponentModel.IContainer components;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox gpbCreateUser;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.Label lblNamePrefix;
        private System.Windows.Forms.Label lblIndex;
        private System.Windows.Forms.TextBox txtNamePrefix;
        private System.Windows.Forms.NumericUpDown nudIndex;
        private System.Windows.Forms.Label lblExample;
        private System.Windows.Forms.Label lblDisplay;
        private System.Windows.Forms.LinkLabel lnkLocation;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.LinkLabel lnkAdduser;
        private System.Windows.Forms.TextBox txtAdduser;
        private System.Windows.Forms.GroupBox gpbDomino;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.TextBox txtServer;
        private System.Windows.Forms.TextBox txtDomain;
        private System.Windows.Forms.NumericUpDown nudNumUser;
        private System.Windows.Forms.Label lblLastName;
        private System.Windows.Forms.Label lblServer;
        private System.Windows.Forms.Label lblDomain;
        private System.Windows.Forms.Label lblShowResult;
        private System.Windows.Forms.Button btnMakeFile;
        private System.Windows.Forms.TextBox txtOutFile;
        private System.Windows.Forms.LinkLabel lnkFolder;
        private System.Windows.Forms.TextBox txtLastPrefix;
        private System.Windows.Forms.Button btnRandNum;
        private System.Windows.Forms.NumericUpDown nudRandNum;
        private System.Windows.Forms.Button btnSeqNum;

        private QATool.CommObj    commObj = new CommObj();

		private string strFolder;

		public Template()
		{
            Debug.WriteLine("Template.cs - Initialize Template Object");
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
                Debug.WriteLine("Template.cs - Deposing Template Object");
				if(components != null)
				{
					Debug.WriteLine("\t Dispose component");
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
            this.lblRcvdEnd = new System.Windows.Forms.Label();
            this.lblNumTemplate = new System.Windows.Forms.Label();
            this.btnModify = new System.Windows.Forms.Button();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.nudCount = new System.Windows.Forms.NumericUpDown();
            this.dtpSendStart = new System.Windows.Forms.DateTimePicker();
            this.txtRepositID = new System.Windows.Forms.TextBox();
            this.dtpSendEnd = new System.Windows.Forms.DateTimePicker();
            this.cboDocType = new System.Windows.Forms.ComboBox();
            this.gbxCriteria1 = new System.Windows.Forms.GroupBox();
            this.label13 = new System.Windows.Forms.Label();
            this.lblBox12 = new System.Windows.Forms.Label();
            this.lblBox11 = new System.Windows.Forms.Label();
            this.txtQ2 = new System.Windows.Forms.TextBox();
            this.txtTag2 = new System.Windows.Forms.TextBox();
            this.txtTag1 = new System.Windows.Forms.TextBox();
            this.txtQ1 = new System.Windows.Forms.TextBox();
            this.lblBox14 = new System.Windows.Forms.Label();
            this.dtpRcvdStart = new System.Windows.Forms.DateTimePicker();
            this.dtpRcvdEnd = new System.Windows.Forms.DateTimePicker();
            this.lblRcvdStart = new System.Windows.Forms.Label();
            this.ttpAudit = new System.Windows.Forms.ToolTip(this.components);
            this.btnCreate = new System.Windows.Forms.Button();
            this.txtQ4 = new System.Windows.Forms.TextBox();
            this.txtQ3 = new System.Windows.Forms.TextBox();
            this.txtTag4 = new System.Windows.Forms.TextBox();
            this.txtTag3 = new System.Windows.Forms.TextBox();
            this.lblPrefix = new System.Windows.Forms.Label();
            this.btnFolder = new System.Windows.Forms.Button();
            this.gpbCreateUser = new System.Windows.Forms.GroupBox();
            this.txtAdduser = new System.Windows.Forms.TextBox();
            this.lnkAdduser = new System.Windows.Forms.LinkLabel();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.lnkLocation = new System.Windows.Forms.LinkLabel();
            this.lblDisplay = new System.Windows.Forms.Label();
            this.lblExample = new System.Windows.Forms.Label();
            this.nudIndex = new System.Windows.Forms.NumericUpDown();
            this.txtNamePrefix = new System.Windows.Forms.TextBox();
            this.lblIndex = new System.Windows.Forms.Label();
            this.lblNamePrefix = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.gpbDomino = new System.Windows.Forms.GroupBox();
            this.lblShowResult = new System.Windows.Forms.Label();
            this.lblLastName = new System.Windows.Forms.Label();
            this.txtLastPrefix = new System.Windows.Forms.TextBox();
            this.lblHeader = new System.Windows.Forms.Label();
            this.nudNumUser = new System.Windows.Forms.NumericUpDown();
            this.btnMakeFile = new System.Windows.Forms.Button();
            this.txtOutFile = new System.Windows.Forms.TextBox();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.txtDomain = new System.Windows.Forms.TextBox();
            this.lblDomain = new System.Windows.Forms.Label();
            this.lblServer = new System.Windows.Forms.Label();
            this.gbxCriteria2 = new System.Windows.Forms.GroupBox();
            this.lblBox21 = new System.Windows.Forms.Label();
            this.lblBox23 = new System.Windows.Forms.Label();
            this.lblBox22 = new System.Windows.Forms.Label();
            this.lblBox24 = new System.Windows.Forms.Label();
            this.lblRepositID = new System.Windows.Forms.Label();
            this.lblStartDate = new System.Windows.Forms.Label();
            this.lblEndDate = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRandNum = new System.Windows.Forms.Button();
            this.nudRandNum = new System.Windows.Forms.NumericUpDown();
            this.btnSeqNum = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nudCount)).BeginInit();
            this.gbxCriteria1.SuspendLayout();
            this.gpbCreateUser.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudIndex)).BeginInit();
            this.gpbDomino.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudNumUser)).BeginInit();
            this.gbxCriteria2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudRandNum)).BeginInit();
            this.SuspendLayout();
            // 
            // lblRcvdEnd
            // 
            this.lblRcvdEnd.Enabled = false;
            this.lblRcvdEnd.Location = new System.Drawing.Point(8, 84);
            this.lblRcvdEnd.Name = "lblRcvdEnd";
            this.lblRcvdEnd.Size = new System.Drawing.Size(120, 16);
            this.lblRcvdEnd.TabIndex = 52;
            this.lblRcvdEnd.Text = "Received End Date = ";
            this.lblRcvdEnd.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblNumTemplate
            // 
            this.lblNumTemplate.Enabled = false;
            this.lblNumTemplate.Location = new System.Drawing.Point(288, 8);
            this.lblNumTemplate.Name = "lblNumTemplate";
            this.lblNumTemplate.Size = new System.Drawing.Size(36, 12);
            this.lblNumTemplate.TabIndex = 47;
            this.lblNumTemplate.Text = "Count";
            this.lblNumTemplate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpAudit.SetToolTip(this.lblNumTemplate, "Number of Template Create");
            // 
            // btnModify
            // 
            this.btnModify.Enabled = false;
            this.btnModify.Location = new System.Drawing.Point(292, 96);
            this.btnModify.Name = "btnModify";
            this.btnModify.Size = new System.Drawing.Size(80, 20);
            this.btnModify.TabIndex = 50;
            this.btnModify.Text = "Modify";
            this.ttpAudit.SetToolTip(this.btnModify, "Modify Date or Repositories ID");
            // 
            // txtPrefix
            // 
            this.txtPrefix.Enabled = false;
            this.txtPrefix.Location = new System.Drawing.Point(288, 48);
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.Size = new System.Drawing.Size(80, 20);
            this.txtPrefix.TabIndex = 49;
            this.txtPrefix.Text = "root";
            this.ttpAudit.SetToolTip(this.txtPrefix, "Criteria file prefix");
            // 
            // nudCount
            // 
            this.nudCount.Enabled = false;
            this.nudCount.Location = new System.Drawing.Point(324, 4);
            this.nudCount.Name = "nudCount";
            this.nudCount.Size = new System.Drawing.Size(48, 20);
            this.nudCount.TabIndex = 46;
            this.ttpAudit.SetToolTip(this.nudCount, "Number of Template");
            this.nudCount.Value = new System.Decimal(new int[] {
                                                                   1,
                                                                   0,
                                                                   0,
                                                                   0});
            // 
            // dtpSendStart
            // 
            this.dtpSendStart.CustomFormat = "MM/dd/yyyy hh:mm tt";
            this.dtpSendStart.Enabled = false;
            this.dtpSendStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpSendStart.Location = new System.Drawing.Point(132, 4);
            this.dtpSendStart.Name = "dtpSendStart";
            this.dtpSendStart.Size = new System.Drawing.Size(148, 20);
            this.dtpSendStart.TabIndex = 38;
            this.ttpAudit.SetToolTip(this.dtpSendStart, "Start Date");
            // 
            // txtRepositID
            // 
            this.txtRepositID.Enabled = false;
            this.txtRepositID.Location = new System.Drawing.Point(132, 124);
            this.txtRepositID.Name = "txtRepositID";
            this.txtRepositID.Size = new System.Drawing.Size(148, 20);
            this.txtRepositID.TabIndex = 42;
            this.txtRepositID.Text = "R0000001";
            this.ttpAudit.SetToolTip(this.txtRepositID, "Type in Repositories ID");
            // 
            // dtpSendEnd
            // 
            this.dtpSendEnd.CustomFormat = "MM/dd/yyyy hh:mm tt";
            this.dtpSendEnd.Enabled = false;
            this.dtpSendEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpSendEnd.Location = new System.Drawing.Point(132, 28);
            this.dtpSendEnd.Name = "dtpSendEnd";
            this.dtpSendEnd.Size = new System.Drawing.Size(148, 20);
            this.dtpSendEnd.TabIndex = 39;
            this.ttpAudit.SetToolTip(this.dtpSendEnd, "End Date");
            // 
            // cboDocType
            // 
            this.cboDocType.Enabled = false;
            this.cboDocType.Items.AddRange(new object[] {
                                                            "INVOICE",
                                                            "=UNINDEXABLE_DOC="});
            this.cboDocType.Location = new System.Drawing.Point(8, 104);
            this.cboDocType.Name = "cboDocType";
            this.cboDocType.Size = new System.Drawing.Size(124, 21);
            this.cboDocType.TabIndex = 40;
            this.cboDocType.Text = "=EMAIL=";
            this.ttpAudit.SetToolTip(this.cboDocType, "Doc type");
            // 
            // gbxCriteria1
            // 
            this.gbxCriteria1.Controls.Add(this.label13);
            this.gbxCriteria1.Controls.Add(this.lblBox12);
            this.gbxCriteria1.Controls.Add(this.lblBox11);
            this.gbxCriteria1.Controls.Add(this.txtQ2);
            this.gbxCriteria1.Controls.Add(this.txtTag2);
            this.gbxCriteria1.Controls.Add(this.txtTag1);
            this.gbxCriteria1.Controls.Add(this.txtQ1);
            this.gbxCriteria1.Controls.Add(this.lblBox14);
            this.gbxCriteria1.Enabled = false;
            this.gbxCriteria1.Location = new System.Drawing.Point(8, 148);
            this.gbxCriteria1.Name = "gbxCriteria1";
            this.gbxCriteria1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.gbxCriteria1.Size = new System.Drawing.Size(176, 64);
            this.gbxCriteria1.TabIndex = 43;
            this.gbxCriteria1.TabStop = false;
            this.gbxCriteria1.Text = "Criteria 1";
            // 
            // label13
            // 
            this.label13.Location = new System.Drawing.Point(4, 44);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(40, 16);
            this.label13.TabIndex = 8;
            this.label13.Text = "Box 13";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBox12
            // 
            this.lblBox12.Location = new System.Drawing.Point(68, 20);
            this.lblBox12.Name = "lblBox12";
            this.lblBox12.Size = new System.Drawing.Size(40, 16);
            this.lblBox12.TabIndex = 7;
            this.lblBox12.Text = "Box 12";
            this.lblBox12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBox11
            // 
            this.lblBox11.Location = new System.Drawing.Point(4, 20);
            this.lblBox11.Name = "lblBox11";
            this.lblBox11.Size = new System.Drawing.Size(40, 16);
            this.lblBox11.TabIndex = 6;
            this.lblBox11.Text = "Box 11";
            this.lblBox11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtQ2
            // 
            this.txtQ2.Location = new System.Drawing.Point(108, 40);
            this.txtQ2.Name = "txtQ2";
            this.txtQ2.Size = new System.Drawing.Size(64, 20);
            this.txtQ2.TabIndex = 5;
            this.txtQ2.Text = "";
            this.ttpAudit.SetToolTip(this.txtQ2, "Query String 14");
            // 
            // txtTag2
            // 
            this.txtTag2.Location = new System.Drawing.Point(44, 40);
            this.txtTag2.Name = "txtTag2";
            this.txtTag2.Size = new System.Drawing.Size(24, 20);
            this.txtTag2.TabIndex = 1;
            this.txtTag2.Text = "";
            this.ttpAudit.SetToolTip(this.txtTag2, "Tag Number 13");
            // 
            // txtTag1
            // 
            this.txtTag1.Location = new System.Drawing.Point(44, 16);
            this.txtTag1.Name = "txtTag1";
            this.txtTag1.Size = new System.Drawing.Size(24, 20);
            this.txtTag1.TabIndex = 0;
            this.txtTag1.Text = "";
            this.ttpAudit.SetToolTip(this.txtTag1, "Tag Number - Box 11");
            // 
            // txtQ1
            // 
            this.txtQ1.Location = new System.Drawing.Point(108, 16);
            this.txtQ1.Name = "txtQ1";
            this.txtQ1.Size = new System.Drawing.Size(64, 20);
            this.txtQ1.TabIndex = 4;
            this.txtQ1.Text = "";
            this.ttpAudit.SetToolTip(this.txtQ1, "Query String - Box 12");
            // 
            // lblBox14
            // 
            this.lblBox14.Location = new System.Drawing.Point(68, 44);
            this.lblBox14.Name = "lblBox14";
            this.lblBox14.Size = new System.Drawing.Size(40, 16);
            this.lblBox14.TabIndex = 55;
            this.lblBox14.Text = "Box 14";
            this.lblBox14.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtpRcvdStart
            // 
            this.dtpRcvdStart.CustomFormat = "MM/dd/yyyy hh:mm tt";
            this.dtpRcvdStart.Enabled = false;
            this.dtpRcvdStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRcvdStart.Location = new System.Drawing.Point(132, 56);
            this.dtpRcvdStart.Name = "dtpRcvdStart";
            this.dtpRcvdStart.Size = new System.Drawing.Size(148, 20);
            this.dtpRcvdStart.TabIndex = 53;
            this.ttpAudit.SetToolTip(this.dtpRcvdStart, "Start Date");
            // 
            // dtpRcvdEnd
            // 
            this.dtpRcvdEnd.CustomFormat = "MM/dd/yyyy hh:mm tt";
            this.dtpRcvdEnd.Enabled = false;
            this.dtpRcvdEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRcvdEnd.Location = new System.Drawing.Point(132, 80);
            this.dtpRcvdEnd.Name = "dtpRcvdEnd";
            this.dtpRcvdEnd.Size = new System.Drawing.Size(148, 20);
            this.dtpRcvdEnd.TabIndex = 54;
            this.ttpAudit.SetToolTip(this.dtpRcvdEnd, "End Date");
            // 
            // lblRcvdStart
            // 
            this.lblRcvdStart.Enabled = false;
            this.lblRcvdStart.Location = new System.Drawing.Point(8, 60);
            this.lblRcvdStart.Name = "lblRcvdStart";
            this.lblRcvdStart.Size = new System.Drawing.Size(120, 16);
            this.lblRcvdStart.TabIndex = 51;
            this.lblRcvdStart.Text = "Received Start Date = ";
            this.lblRcvdStart.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnCreate
            // 
            this.btnCreate.Enabled = false;
            this.btnCreate.Location = new System.Drawing.Point(292, 72);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(80, 20);
            this.btnCreate.TabIndex = 48;
            this.btnCreate.Text = "Create";
            this.ttpAudit.SetToolTip(this.btnCreate, "Create audit templat");
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // txtQ4
            // 
            this.txtQ4.Location = new System.Drawing.Point(116, 40);
            this.txtQ4.Name = "txtQ4";
            this.txtQ4.Size = new System.Drawing.Size(64, 20);
            this.txtQ4.TabIndex = 6;
            this.txtQ4.Text = "";
            this.ttpAudit.SetToolTip(this.txtQ4, "Query String - box 24");
            // 
            // txtQ3
            // 
            this.txtQ3.Location = new System.Drawing.Point(116, 16);
            this.txtQ3.Name = "txtQ3";
            this.txtQ3.Size = new System.Drawing.Size(64, 20);
            this.txtQ3.TabIndex = 5;
            this.txtQ3.Text = "";
            this.ttpAudit.SetToolTip(this.txtQ3, "Query String - box 22");
            // 
            // txtTag4
            // 
            this.txtTag4.Location = new System.Drawing.Point(48, 40);
            this.txtTag4.Name = "txtTag4";
            this.txtTag4.Size = new System.Drawing.Size(24, 20);
            this.txtTag4.TabIndex = 3;
            this.txtTag4.Text = "";
            this.ttpAudit.SetToolTip(this.txtTag4, "Tag Number - box 23");
            // 
            // txtTag3
            // 
            this.txtTag3.Location = new System.Drawing.Point(48, 16);
            this.txtTag3.Name = "txtTag3";
            this.txtTag3.Size = new System.Drawing.Size(24, 20);
            this.txtTag3.TabIndex = 2;
            this.txtTag3.Text = "";
            this.ttpAudit.SetToolTip(this.txtTag3, "Tag Number - box 21");
            // 
            // lblPrefix
            // 
            this.lblPrefix.Enabled = false;
            this.lblPrefix.Location = new System.Drawing.Point(288, 28);
            this.lblPrefix.Name = "lblPrefix";
            this.lblPrefix.Size = new System.Drawing.Size(56, 16);
            this.lblPrefix.TabIndex = 55;
            this.lblPrefix.Text = "File Prefix";
            this.lblPrefix.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpAudit.SetToolTip(this.lblPrefix, "Number of Template Create");
            // 
            // btnFolder
            // 
            this.btnFolder.Enabled = false;
            this.btnFolder.Location = new System.Drawing.Point(292, 124);
            this.btnFolder.Name = "btnFolder";
            this.btnFolder.Size = new System.Drawing.Size(80, 20);
            this.btnFolder.TabIndex = 57;
            this.btnFolder.Text = "Folder";
            this.ttpAudit.SetToolTip(this.btnFolder, "Criteria files Folder");
            this.btnFolder.Click += new System.EventHandler(this.btnFolder_Click);
            // 
            // gpbCreateUser
            // 
            this.gpbCreateUser.Controls.Add(this.btnSeqNum);
            this.gpbCreateUser.Controls.Add(this.nudRandNum);
            this.gpbCreateUser.Controls.Add(this.btnRandNum);
            this.gpbCreateUser.Controls.Add(this.txtAdduser);
            this.gpbCreateUser.Controls.Add(this.lnkAdduser);
            this.gpbCreateUser.Controls.Add(this.txtFileName);
            this.gpbCreateUser.Controls.Add(this.lnkLocation);
            this.gpbCreateUser.Controls.Add(this.lblDisplay);
            this.gpbCreateUser.Controls.Add(this.lblExample);
            this.gpbCreateUser.Controls.Add(this.nudIndex);
            this.gpbCreateUser.Controls.Add(this.txtNamePrefix);
            this.gpbCreateUser.Controls.Add(this.lblIndex);
            this.gpbCreateUser.Controls.Add(this.lblNamePrefix);
            this.gpbCreateUser.Controls.Add(this.lblUserName);
            this.gpbCreateUser.Controls.Add(this.btnGenerate);
            this.gpbCreateUser.Location = new System.Drawing.Point(8, 216);
            this.gpbCreateUser.Name = "gpbCreateUser";
            this.gpbCreateUser.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.gpbCreateUser.Size = new System.Drawing.Size(368, 92);
            this.gpbCreateUser.TabIndex = 60;
            this.gpbCreateUser.TabStop = false;
            this.gpbCreateUser.Text = "Create list of users for AD";
            this.ttpAudit.SetToolTip(this.gpbCreateUser, "Create an input file for addusers.exe");
            // 
            // txtAdduser
            // 
            this.txtAdduser.Location = new System.Drawing.Point(240, 40);
            this.txtAdduser.Name = "txtAdduser";
            this.txtAdduser.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtAdduser.Size = new System.Drawing.Size(124, 20);
            this.txtAdduser.TabIndex = 63;
            this.txtAdduser.Text = "point to adduser.exe";
            // 
            // lnkAdduser
            // 
            this.lnkAdduser.Location = new System.Drawing.Point(168, 44);
            this.lnkAdduser.Name = "lnkAdduser";
            this.lnkAdduser.Size = new System.Drawing.Size(68, 20);
            this.lnkAdduser.TabIndex = 62;
            this.lnkAdduser.TabStop = true;
            this.lnkAdduser.Text = "AddUsers";
            this.lnkAdduser.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkAdduser.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkAdduser_LinkClicked);
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(240, 20);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtFileName.Size = new System.Drawing.Size(124, 20);
            this.txtFileName.TabIndex = 8;
            this.txtFileName.Text = "c:\\tmp";
            this.ttpAudit.SetToolTip(this.txtFileName, "The output file name");
            // 
            // lnkLocation
            // 
            this.lnkLocation.Location = new System.Drawing.Point(168, 24);
            this.lnkLocation.Name = "lnkLocation";
            this.lnkLocation.Size = new System.Drawing.Size(68, 20);
            this.lnkLocation.TabIndex = 7;
            this.lnkLocation.TabStop = true;
            this.lnkLocation.Text = "File Location";
            this.lnkLocation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLocation.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkLocation_LinkClicked);
            // 
            // lblDisplay
            // 
            this.lblDisplay.Location = new System.Drawing.Point(68, 44);
            this.lblDisplay.Name = "lblDisplay";
            this.lblDisplay.Size = new System.Drawing.Size(100, 16);
            this.lblDisplay.TabIndex = 6;
            this.lblDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ttpAudit.SetToolTip(this.lblDisplay, "User name in AD");
            // 
            // lblExample
            // 
            this.lblExample.Location = new System.Drawing.Point(16, 44);
            this.lblExample.Name = "lblExample";
            this.lblExample.Size = new System.Drawing.Size(48, 16);
            this.lblExample.TabIndex = 5;
            this.lblExample.Text = "Example";
            this.lblExample.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // nudIndex
            // 
            this.nudIndex.Cursor = System.Windows.Forms.Cursors.Default;
            this.nudIndex.Increment = new System.Decimal(new int[] {
                                                                       10,
                                                                       0,
                                                                       0,
                                                                       0});
            this.nudIndex.Location = new System.Drawing.Point(116, 24);
            this.nudIndex.Maximum = new System.Decimal(new int[] {
                                                                     5000,
                                                                     0,
                                                                     0,
                                                                     0});
            this.nudIndex.Name = "nudIndex";
            this.nudIndex.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nudIndex.Size = new System.Drawing.Size(52, 20);
            this.nudIndex.TabIndex = 4;
            this.ttpAudit.SetToolTip(this.nudIndex, "0 .. 5000");
            this.nudIndex.Value = new System.Decimal(new int[] {
                                                                   2000,
                                                                   0,
                                                                   0,
                                                                   0});
            this.nudIndex.KeyUp += new System.Windows.Forms.KeyEventHandler(this.nudIndex_KeyUp);
            this.nudIndex.ValueChanged += new System.EventHandler(this.nudIndex_ValueChanged);
            // 
            // txtNamePrefix
            // 
            this.txtNamePrefix.Location = new System.Drawing.Point(64, 24);
            this.txtNamePrefix.Name = "txtNamePrefix";
            this.txtNamePrefix.Size = new System.Drawing.Size(52, 20);
            this.txtNamePrefix.TabIndex = 3;
            this.txtNamePrefix.Text = "zuA";
            this.ttpAudit.SetToolTip(this.txtNamePrefix, "Prefix");
            this.txtNamePrefix.TextChanged += new System.EventHandler(this.txtNamePrefix_TextChanged);
            // 
            // lblIndex
            // 
            this.lblIndex.Location = new System.Drawing.Point(124, 8);
            this.lblIndex.Name = "lblIndex";
            this.lblIndex.Size = new System.Drawing.Size(36, 16);
            this.lblIndex.TabIndex = 2;
            this.lblIndex.Text = "Index";
            this.lblIndex.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ttpAudit.SetToolTip(this.lblIndex, "Second part of the name");
            // 
            // lblNamePrefix
            // 
            this.lblNamePrefix.Location = new System.Drawing.Point(72, 8);
            this.lblNamePrefix.Name = "lblNamePrefix";
            this.lblNamePrefix.Size = new System.Drawing.Size(36, 16);
            this.lblNamePrefix.TabIndex = 1;
            this.lblNamePrefix.Text = "Prefix";
            this.lblNamePrefix.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ttpAudit.SetToolTip(this.lblNamePrefix, "First part of the name");
            // 
            // lblUserName
            // 
            this.lblUserName.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblUserName.Location = new System.Drawing.Point(4, 28);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(64, 16);
            this.lblUserName.TabIndex = 0;
            this.lblUserName.Text = "User Name";
            this.lblUserName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpAudit.SetToolTip(this.lblUserName, "Contain 2 parts");
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(280, 64);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(80, 20);
            this.btnGenerate.TabIndex = 61;
            this.btnGenerate.Text = "Generate";
            this.ttpAudit.SetToolTip(this.btnGenerate, "Execute addusers.exe in shell");
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // gpbDomino
            // 
            this.gpbDomino.Controls.Add(this.lblShowResult);
            this.gpbDomino.Controls.Add(this.lblLastName);
            this.gpbDomino.Controls.Add(this.txtLastPrefix);
            this.gpbDomino.Controls.Add(this.lblHeader);
            this.gpbDomino.Controls.Add(this.nudNumUser);
            this.gpbDomino.Controls.Add(this.btnMakeFile);
            this.gpbDomino.Controls.Add(this.txtOutFile);
            this.gpbDomino.Controls.Add(this.lnkFolder);
            this.gpbDomino.Controls.Add(this.txtServer);
            this.gpbDomino.Controls.Add(this.txtDomain);
            this.gpbDomino.Controls.Add(this.lblDomain);
            this.gpbDomino.Controls.Add(this.lblServer);
            this.gpbDomino.Location = new System.Drawing.Point(8, 312);
            this.gpbDomino.Name = "gpbDomino";
            this.gpbDomino.Size = new System.Drawing.Size(368, 108);
            this.gpbDomino.TabIndex = 61;
            this.gpbDomino.TabStop = false;
            this.ttpAudit.SetToolTip(this.gpbDomino, "Create list of user for domino");
            // 
            // lblShowResult
            // 
            this.lblShowResult.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblShowResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblShowResult.Location = new System.Drawing.Point(4, 80);
            this.lblShowResult.Name = "lblShowResult";
            this.lblShowResult.Size = new System.Drawing.Size(360, 20);
            this.lblShowResult.TabIndex = 69;
            this.ttpAudit.SetToolTip(this.lblShowResult, "Display the result");
            // 
            // lblLastName
            // 
            this.lblLastName.Location = new System.Drawing.Point(4, 36);
            this.lblLastName.Name = "lblLastName";
            this.lblLastName.Size = new System.Drawing.Size(60, 12);
            this.lblLastName.TabIndex = 65;
            this.lblLastName.Text = "Last Name";
            // 
            // txtLastPrefix
            // 
            this.txtLastPrefix.Location = new System.Drawing.Point(64, 32);
            this.txtLastPrefix.Name = "txtLastPrefix";
            this.txtLastPrefix.Size = new System.Drawing.Size(20, 20);
            this.txtLastPrefix.TabIndex = 1;
            this.txtLastPrefix.Text = "A";
            this.txtLastPrefix.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.ttpAudit.SetToolTip(this.txtLastPrefix, "Last Name Prefix");
            this.txtLastPrefix.TextChanged += new System.EventHandler(this.txtLastPrefix_TextChanged);
            // 
            // lblHeader
            // 
            this.lblHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblHeader.Location = new System.Drawing.Point(8, 12);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Size = new System.Drawing.Size(360, 16);
            this.lblHeader.TabIndex = 0;
            this.lblHeader.Text = "A0011; userA; A; ; passwd0; ; ; ServerName / domcert; ; ; ; ; ; user Profile";
            // 
            // nudNumUser
            // 
            this.nudNumUser.Cursor = System.Windows.Forms.Cursors.Default;
            this.nudNumUser.Increment = new System.Decimal(new int[] {
                                                                         10,
                                                                         0,
                                                                         0,
                                                                         0});
            this.nudNumUser.Location = new System.Drawing.Point(84, 32);
            this.nudNumUser.Maximum = new System.Decimal(new int[] {
                                                                       5000,
                                                                       0,
                                                                       0,
                                                                       0});
            this.nudNumUser.Name = "nudNumUser";
            this.nudNumUser.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nudNumUser.Size = new System.Drawing.Size(52, 20);
            this.nudNumUser.TabIndex = 64;
            this.nudNumUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.ttpAudit.SetToolTip(this.nudNumUser, "0 .. 5000");
            this.nudNumUser.Value = new System.Decimal(new int[] {
                                                                     2020,
                                                                     0,
                                                                     0,
                                                                     0});
            this.nudNumUser.KeyUp += new System.Windows.Forms.KeyEventHandler(this.nudNumUser_KeyUp);
            this.nudNumUser.ValueChanged += new System.EventHandler(this.txtNamePrefix_TextChanged);
            // 
            // btnMakeFile
            // 
            this.btnMakeFile.Location = new System.Drawing.Point(280, 56);
            this.btnMakeFile.Name = "btnMakeFile";
            this.btnMakeFile.Size = new System.Drawing.Size(80, 20);
            this.btnMakeFile.TabIndex = 64;
            this.btnMakeFile.Text = "Make";
            this.ttpAudit.SetToolTip(this.btnMakeFile, "Make the import text for Domino");
            this.btnMakeFile.Click += new System.EventHandler(this.btnMakeFile_Click);
            // 
            // txtOutFile
            // 
            this.txtOutFile.Location = new System.Drawing.Point(184, 32);
            this.txtOutFile.Name = "txtOutFile";
            this.txtOutFile.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtOutFile.Size = new System.Drawing.Size(176, 20);
            this.txtOutFile.TabIndex = 65;
            this.txtOutFile.Text = "c:\\tmp";
            this.ttpAudit.SetToolTip(this.txtOutFile, "Output File Name");
            // 
            // lnkFolder
            // 
            this.lnkFolder.Location = new System.Drawing.Point(144, 32);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(36, 20);
            this.lnkFolder.TabIndex = 64;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(64, 56);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(70, 20);
            this.txtServer.TabIndex = 2;
            this.txtServer.Text = "R5Dom";
            this.ttpAudit.SetToolTip(this.txtServer, "Server Name");
            this.txtServer.TextChanged += new System.EventHandler(this.txtServer_TextChanged);
            // 
            // txtDomain
            // 
            this.txtDomain.Location = new System.Drawing.Point(184, 56);
            this.txtDomain.Name = "txtDomain";
            this.txtDomain.Size = new System.Drawing.Size(92, 20);
            this.txtDomain.TabIndex = 3;
            this.txtDomain.Text = "OrgName";
            this.ttpAudit.SetToolTip(this.txtDomain, "Domino Domain Name");
            this.txtDomain.TextChanged += new System.EventHandler(this.txtDomain_TextChanged);
            // 
            // lblDomain
            // 
            this.lblDomain.Location = new System.Drawing.Point(136, 60);
            this.lblDomain.Name = "lblDomain";
            this.lblDomain.Size = new System.Drawing.Size(44, 12);
            this.lblDomain.TabIndex = 67;
            this.lblDomain.Text = "Org ID";
            this.ttpAudit.SetToolTip(this.lblDomain, "Domino Organization Name");
            // 
            // lblServer
            // 
            this.lblServer.Location = new System.Drawing.Point(24, 60);
            this.lblServer.Name = "lblServer";
            this.lblServer.Size = new System.Drawing.Size(40, 12);
            this.lblServer.TabIndex = 66;
            this.lblServer.Text = "Server";
            this.ttpAudit.SetToolTip(this.lblServer, "Domino Server Name");
            // 
            // gbxCriteria2
            // 
            this.gbxCriteria2.Controls.Add(this.lblBox21);
            this.gbxCriteria2.Controls.Add(this.txtQ4);
            this.gbxCriteria2.Controls.Add(this.txtQ3);
            this.gbxCriteria2.Controls.Add(this.txtTag4);
            this.gbxCriteria2.Controls.Add(this.txtTag3);
            this.gbxCriteria2.Controls.Add(this.lblBox23);
            this.gbxCriteria2.Controls.Add(this.lblBox22);
            this.gbxCriteria2.Controls.Add(this.lblBox24);
            this.gbxCriteria2.Enabled = false;
            this.gbxCriteria2.Location = new System.Drawing.Point(188, 148);
            this.gbxCriteria2.Name = "gbxCriteria2";
            this.gbxCriteria2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.gbxCriteria2.Size = new System.Drawing.Size(188, 64);
            this.gbxCriteria2.TabIndex = 44;
            this.gbxCriteria2.TabStop = false;
            this.gbxCriteria2.Text = "Criteria 2";
            // 
            // lblBox21
            // 
            this.lblBox21.Location = new System.Drawing.Point(4, 20);
            this.lblBox21.Name = "lblBox21";
            this.lblBox21.Size = new System.Drawing.Size(40, 16);
            this.lblBox21.TabIndex = 7;
            this.lblBox21.Text = "Box 21";
            this.lblBox21.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBox23
            // 
            this.lblBox23.Location = new System.Drawing.Point(4, 44);
            this.lblBox23.Name = "lblBox23";
            this.lblBox23.Size = new System.Drawing.Size(40, 16);
            this.lblBox23.TabIndex = 55;
            this.lblBox23.Text = "Box 23";
            this.lblBox23.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBox22
            // 
            this.lblBox22.Location = new System.Drawing.Point(76, 20);
            this.lblBox22.Name = "lblBox22";
            this.lblBox22.Size = new System.Drawing.Size(40, 16);
            this.lblBox22.TabIndex = 55;
            this.lblBox22.Text = "Box 22";
            this.lblBox22.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBox24
            // 
            this.lblBox24.Location = new System.Drawing.Point(76, 44);
            this.lblBox24.Name = "lblBox24";
            this.lblBox24.Size = new System.Drawing.Size(40, 16);
            this.lblBox24.TabIndex = 55;
            this.lblBox24.Text = "Box 24";
            this.lblBox24.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblRepositID
            // 
            this.lblRepositID.Enabled = false;
            this.lblRepositID.Location = new System.Drawing.Point(8, 128);
            this.lblRepositID.Name = "lblRepositID";
            this.lblRepositID.Size = new System.Drawing.Size(94, 16);
            this.lblRepositID.TabIndex = 41;
            this.lblRepositID.Text = "REPOSITORIES:";
            this.lblRepositID.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblStartDate
            // 
            this.lblStartDate.Enabled = false;
            this.lblStartDate.Location = new System.Drawing.Point(8, 8);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(120, 16);
            this.lblStartDate.TabIndex = 36;
            this.lblStartDate.Text = "Sent Start Date = ";
            this.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblEndDate
            // 
            this.lblEndDate.Enabled = false;
            this.lblEndDate.Location = new System.Drawing.Point(8, 32);
            this.lblEndDate.Name = "lblEndDate";
            this.lblEndDate.Size = new System.Drawing.Size(120, 16);
            this.lblEndDate.TabIndex = 37;
            this.lblEndDate.Text = "Sent End Date = ";
            this.lblEndDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Yellow;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.label1.Location = new System.Drawing.Point(136, 100);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 20);
            this.label1.TabIndex = 59;
            this.label1.Text = "R e t i r e d";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnRandNum
            // 
            this.btnRandNum.Location = new System.Drawing.Point(192, 64);
            this.btnRandNum.Name = "btnRandNum";
            this.btnRandNum.Size = new System.Drawing.Size(80, 20);
            this.btnRandNum.TabIndex = 64;
            this.btnRandNum.Text = "Gen Rand";
            this.ttpAudit.SetToolTip(this.btnRandNum, "Generate Random number");
            this.btnRandNum.Click += new System.EventHandler(this.btnRandNum_Click);
            // 
            // nudRandNum
            // 
            this.nudRandNum.Cursor = System.Windows.Forms.Cursors.Default;
            this.nudRandNum.Increment = new System.Decimal(new int[] {
                                                                         10,
                                                                         0,
                                                                         0,
                                                                         0});
            this.nudRandNum.Location = new System.Drawing.Point(8, 64);
            this.nudRandNum.Maximum = new System.Decimal(new int[] {
                                                                       50000000,
                                                                       0,
                                                                       0,
                                                                       0});
            this.nudRandNum.Name = "nudRandNum";
            this.nudRandNum.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nudRandNum.Size = new System.Drawing.Size(88, 20);
            this.nudRandNum.TabIndex = 65;
            this.ttpAudit.SetToolTip(this.nudRandNum, "0 .. 50000000");
            this.nudRandNum.Value = new System.Decimal(new int[] {
                                                                     50000000,
                                                                     0,
                                                                     0,
                                                                     0});
            // 
            // btnSeqNum
            // 
            this.btnSeqNum.Location = new System.Drawing.Point(108, 64);
            this.btnSeqNum.Name = "btnSeqNum";
            this.btnSeqNum.Size = new System.Drawing.Size(80, 20);
            this.btnSeqNum.TabIndex = 66;
            this.btnSeqNum.Text = "Seq Number";
            this.ttpAudit.SetToolTip(this.btnSeqNum, "Generate Random number");
            this.btnSeqNum.Click += new System.EventHandler(this.btnSeqNum_Click);
            // 
            // Template
            // 
            this.Controls.Add(this.gpbDomino);
            this.Controls.Add(this.gpbCreateUser);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnFolder);
            this.Controls.Add(this.lblPrefix);
            this.Controls.Add(this.lblNumTemplate);
            this.Controls.Add(this.btnModify);
            this.Controls.Add(this.txtPrefix);
            this.Controls.Add(this.nudCount);
            this.Controls.Add(this.dtpSendStart);
            this.Controls.Add(this.txtRepositID);
            this.Controls.Add(this.dtpSendEnd);
            this.Controls.Add(this.cboDocType);
            this.Controls.Add(this.gbxCriteria1);
            this.Controls.Add(this.dtpRcvdStart);
            this.Controls.Add(this.dtpRcvdEnd);
            this.Controls.Add(this.lblRcvdStart);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.gbxCriteria2);
            this.Controls.Add(this.lblRepositID);
            this.Controls.Add(this.lblStartDate);
            this.Controls.Add(this.lblEndDate);
            this.Controls.Add(this.lblRcvdEnd);
            this.Name = "Template";
            this.Size = new System.Drawing.Size(388, 424);
            this.Load += new System.EventHandler(this.Template_Load);
            ((System.ComponentModel.ISupportInitialize)(this.nudCount)).EndInit();
            this.gbxCriteria1.ResumeLayout(false);
            this.gpbCreateUser.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudIndex)).EndInit();
            this.gpbDomino.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudNumUser)).EndInit();
            this.gbxCriteria2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.nudRandNum)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

        #region Retired
		private void btnCreate_Click(object sender, System.EventArgs e)
		{
			if( !ValidUserInput() )
				return;

			string line = "";
			string fn = "";
			this.Cursor = Cursors.WaitCursor;

			// Compose a string that consists of three lines.
			for( int i = 0; i < nudCount.Value; i++ )
			{
				fn = txtPrefix.Text + i + ".criteria-parse";
				StreamWriter sw = new StreamWriter( fn );
				try
				{
					// Write the string to a file.				
					line = "BEGSENTDATE=" + dtpSendStart.Text; 
					sw.WriteLine( line ); // Sent start date
					line = "ENDSENTDATE=" + dtpSendEnd.Text;
					sw.WriteLine( line ); // Send end date

					line = "BEGRCVDDATE=" + dtpRcvdStart.Text; 
					sw.WriteLine( line ); // Received start date
					line = "ENDRCVDDATE=" + dtpRcvdEnd.Text;
					sw.WriteLine( line ); // Received end date

					sw.WriteLine( cboDocType.Text );
					line = "REPOSITORIES:" + txtRepositID.Text + " AND";
					sw.WriteLine( line );
					sw.WriteLine( "(" );

					if( txtTag2.TextLength != 0 )
					{
						line = "    (" + txtTag1.Text + ":" + txtQ1.Text + ") OR";
						sw.WriteLine( line );
						line = "    (" + txtTag2.Text + ":" + txtQ2.Text + ")";
						sw.WriteLine( line );
					} // end of txtTag2 != 0
					else
					{
						line = "    (" + txtTag1.Text + ":" + txtQ1.Text + ")";
						sw.WriteLine( line );
					}

					sw.WriteLine( ") AND" ); // another field
					sw.WriteLine("(");
				

					if( txtTag4.TextLength != 0 )
					{
						line = "    (" + txtTag3.Text + ":" + txtQ3.Text + ") OR";
						sw.WriteLine( line );

						line = "    (" + txtTag4.Text + ":" + txtQ4.Text + ")";
						sw.WriteLine( line );
					}//end of txtTag4 != 0
					else
					{
						line = "    (" + txtTag3.Text + ":" + txtQ3.Text + ")";
						sw.WriteLine( line );
					}

					sw.WriteLine(")");
				}//end of try
				catch( IOException ex )
				{
					MessageBox.Show(ex.Message.ToString(), "IOException");
				}// end of catch

				sw.Close();
			}//end of for - count loop
			this.Cursor = Cursors.Default;		
		}//end of btnCreate_Click

		/// <summary>
		/// User input validation
		/// If first number tag text box and Query string text box are empty in both criteria,
		/// Warning message box pop up for user input.
		/// </summary>
		/// <returns>true - OK; false - FAIL</returns>
		public bool ValidUserInput()
		{
			bool rv = true; // initialize as true - no error

			txtTag1.Text = txtTag1.Text.Trim( ' ' ); // remove leading and ending space
			txtTag2.Text = txtTag2.Text.Trim( ' ' );
			txtTag3.Text = txtTag3.Text.Trim( ' ' );
			txtTag4.Text = txtTag4.Text.Trim( ' ' );

			txtQ1.Text = txtQ1.Text.Trim( ' ' );
			txtQ2.Text = txtQ2.Text.Trim( ' ' );
			txtQ3.Text = txtQ3.Text.Trim( ' ' );
			txtQ4.Text = txtQ4.Text.Trim( ' ' );

			if( txtTag1.Text.Length == 0 )
			{
				MessageBox.Show( "Please set the tag number in Box 11", "Warning" );				
				txtTag1.Focus();
				return( false );
			}//end of if - txtTag1

			if( txtQ1.Text.Length == 0 )
			{
				MessageBox.Show( "Please set the Query String in Box 12", "Warning" );
				txtQ1.Focus();
				return( false );
			}//end of if - txtQ1

			if( txtTag3.Text.Length == 0 )
			{
				MessageBox.Show( "Please set the tag number in Box 21", "Warning" );
				txtTag3.Focus();
				return( false );
			}// end of if -txtTag3

			if( txtQ3.Text.Length == 0 )
			{
				MessageBox.Show( "Please set the Query String in Box 22", "Warning" );
				txtQ3.Focus();
				return( false );
			}//end of if - txtQ3

			return( rv );
		}// end of ValidUserInput

		private void btnFolder_Click(object sender, System.EventArgs e)
		{
			FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
			if( fbDlg.ShowDialog() == DialogResult.OK )
			{
				strFolder = fbDlg.SelectedPath;
			}
		}//end of btnFolder_Click
        #endregion

        private void txtNamePrefix_TextChanged(object sender, System.EventArgs e)
        {        
            lblDisplay.Text = txtNamePrefix.Text + nudIndex.Value.ToString().PadLeft(4,'0');
        }//end of txtNamePrefix_TextChanged

        private void nudIndex_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            lblDisplay.Text = txtNamePrefix.Text + nudIndex.Value.ToString().PadLeft(4,'0');
        }//end of nudIndex_KeyUp

        private void nudIndex_ValueChanged(object sender, System.EventArgs e)
        {
            lblDisplay.Text = txtNamePrefix.Text + nudIndex.Value.ToString().PadLeft(4,'0');
        }//end of nudIndex_ValueChanged

        private void Template_Load(object sender, System.EventArgs e)
        {
            lblDisplay.Text = txtNamePrefix.Text + nudIndex.Value.ToString().PadLeft(4,'0');

            DisplayInfo();
        }//end of Template_Load

        private void lnkLocation_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
/*** save for reference : do nothing
            Trace.WriteLine( "Template.cs - lnkFolder_LinkClicked" );
            FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
            if( txtFileName.Text != null )
                fbDlg.SelectedPath = txtFileName.Text;  // set the default folder

            if( fbDlg.ShowDialog() == DialogResult.OK )
            {
                txtFileName.Text = fbDlg.SelectedPath;
            }
 save for reference : do nothing **/
        }//end of lnkLocation_LinkClicked

        private void btnGenerate_Click(object sender, System.EventArgs e)
        {
            string line = "";
            string fn = txtFileName.Text;
            this.Cursor = Cursors.WaitCursor;

            // TO DO : Validate user input of txtAdduser.text
            // TO DO - check do addusers.exe and <user.txt> exist?            

            StreamWriter sw = new StreamWriter( fn );

            try
            {
                sw.WriteLine( "[User]" );

                for( int i = 0; i < nudIndex.Value; i++ )
                {
                    // duserA11,duserA11 groupA1,password0,,,,,
                    line = txtNamePrefix.Text + i.ToString().PadLeft(4,'0') 
                        + "," + txtNamePrefix.Text + i.ToString().PadLeft(4, '0')
                        + " Group" + txtNamePrefix.Text
                        + ",password0,,,,,";                   

                    sw.WriteLine( line );
                }//end of for

                Process.Start( txtAdduser.Text, " /c " + txtFileName.Text );
            }//end of try
            catch( ArgumentException argEx )
            {
                MessageBox.Show(argEx.Message.ToString(), "Argument Exception" );
            }//end of catch - process
            catch( Win32Exception w32Ex )
            {
                MessageBox.Show( w32Ex.Message.ToString(), "Win32 Exception" );
            }//end of catch - process
            catch( IOException ex )
            {
                MessageBox.Show(ex.Message.ToString(), "IOException");
            }// end of catch - stream writer
            sw.Close();

            this.Cursor = Cursors.Default;		        
        }//end of btnGenerate_Click

        private void lnkAdduser_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            // Show the Open File Dialog.
            // If the user clicked OK in the dialog and a text file was selected, open it.
            // Display an OpenFileDialog and show the read-only files...
            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.RestoreDirectory = false;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtAdduser.Text = ofDlg.FileName;
            }//end of if

            txtFileName.Text = Directory.GetCurrentDirectory() + "\\" + txtNamePrefix.Text + ".txt";
        }//end of lnkAdduser_LinkClicked

        private void btnMakeFile_Click(object sender, System.EventArgs e)
        {
            string line = "";
            string fn = txtOutFile.Text;
            this.Cursor = Cursors.WaitCursor;

            // TO DO : Validate user input of txtAdduser.text
            // TO DO - check do addusers.exe and <user.txt> exist?            

            StreamWriter sw = new StreamWriter( fn );

            try
            {
                for( int i = 0; i < nudNumUser.Value; i++ )
                {
                    // A0011;userA;A;;password0;;;w2kDomino / domain;;;;;;user Profile
                    line = txtLastPrefix.Text + i.ToString().PadLeft(4,'0') + ";" 
                        + "user" + txtLastPrefix.Text + ";"
                        + txtLastPrefix.Text + ";;"
                        + "password0;;;"
                        + txtServer.Text + " / "
                        + txtDomain.Text + ";;;;;;"
                        + txtLastPrefix.Text + i.ToString().PadLeft(4,'0') + " Profile";

                    lblShowResult.Text = line;
                    sw.WriteLine( line );
                }//end of for
            }//end of try
            catch( ArgumentException argEx )
            {
                MessageBox.Show(argEx.Message.ToString(), "Argument Exception" );
            }//end of catch - process
            catch( Win32Exception w32Ex )
            {
                MessageBox.Show( w32Ex.Message.ToString(), "Win32 Exception" );
            }//end of catch - process
            catch( IOException ex )
            {
                MessageBox.Show(ex.Message.ToString(), "IOException");
            }// end of catch - stream writer
            sw.Close();

            this.Cursor = Cursors.Default;		                
        }//end of btnMakeFile_Click

        private void DisplayInfo()
        {                   
            string line = txtLastPrefix.Text + nudNumUser.Value.ToString().PadLeft(4,'0') + ";" 
                + "user" + txtLastPrefix.Text + ";"
                + txtLastPrefix.Text + ";;"
                + "password0;;;"
                + txtServer.Text + " / "
                + txtDomain.Text + ";;;;;;"
                + txtLastPrefix.Text + nudNumUser.Value.ToString().PadLeft(4,'0') + " Profile";

            lblShowResult.Text = line;
        }//end of DisplayInfo
        private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            if( fbDlg.ShowDialog() == DialogResult.OK )
            {
                txtOutFile.Text = fbDlg.SelectedPath + "\\" + txtLastPrefix.Text + "User.txt";
            }
        }

        private void nudNumUser_ValueChanged(object sender, System.EventArgs e)
        {
            DisplayInfo();
        }

        private void txtLastPrefix_TextChanged(object sender, System.EventArgs e)
        {
            DisplayInfo();
        }

        private void txtDomain_TextChanged(object sender, System.EventArgs e)
        {
            DisplayInfo();
        }

        private void txtServer_TextChanged(object sender, System.EventArgs e)
        {
            DisplayInfo();
        }

        // Value change event only happen when enter key was hit for up down control
        // Therefore, using key up event
        private void nudNumUser_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            DisplayInfo();
        }

        private void btnRandNum_Click(object sender, System.EventArgs e)
        {
            // Create a random object with a timer-generated seed.
            // Wait to allow the timer to advance.
            Thread.Sleep( 1 );
            Random autoRand = new Random( );

            string str = "";
            string fileName = txtFileName.Text;
            for( int j = 0; j < nudRandNum.Value; j++ )
            {
                str = j.ToString() + ": " + autoRand.Next().ToString();
                commObj.LogGUID( fileName, str );
            }
        }// end of btnRandNum_Click

        private void btnSeqNum_Click(object sender, System.EventArgs e)
        {
            for ( int i = 0; i < nudRandNum.Value; i++ )
            {
                commObj.LogGUID( txtFileName.Text, i.ToString() );
            }//end of for
        
        }
	}
}
