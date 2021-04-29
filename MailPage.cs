using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Web;
using System.Web.Mail;

namespace QATool
{
	/// <summary>
	/// Summary description for MailPage.
	/// </summary>
    public class MailPage : System.Windows.Forms.UserControl
    {
        private System.Windows.Forms.LinkLabel lnkFrom;
        private System.Windows.Forms.LinkLabel lnkTo;
        private System.Windows.Forms.LinkLabel lnkCC;
        private System.Windows.Forms.LinkLabel lnkBCC;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.RichTextBox richBox;
        private System.Windows.Forms.LinkLabel lnkAttach;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Label lblSMTP;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.ComboBox cboSMTP;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.ToolTip ttpMailPage;
        private System.Windows.Forms.TextBox txtAttach;
        private System.ComponentModel.IContainer components;

        // custom variable
        private String msgCaption = "Mail Page";
        private System.Windows.Forms.CheckBox chkHeader;
        private System.Windows.Forms.CheckBox chkConSearch;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.ComboBox cboHeader1;
        private System.Windows.Forms.ComboBox cboHeader2;
        private System.Windows.Forms.ComboBox cboHeaderVal1;
        private System.Windows.Forms.ComboBox cboHeaderVal2;
        private System.Windows.Forms.ComboBox cboTo;
        private System.Windows.Forms.ComboBox cboFrom;
        private System.Windows.Forms.ComboBox cboCC;
        private System.Windows.Forms.ComboBox cboBCC;
        private System.Windows.Forms.CheckBox chkMultiAttach;
        private System.Windows.Forms.RadioButton rdoMailAPI;
        private System.Windows.Forms.RadioButton rdoSMTPClient;
        private System.Windows.Forms.DateTimePicker dtpSendDate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblCustSendDate;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkUnparsable;
        private System.Windows.Forms.CheckBox chkUnindexable; 

        private QATool.CommObj    commObj = new CommObj();

        public MailPage()
        {
            Debug.WriteLine("MailPage.cs - Initialize MailPage Object");

            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();

            // TODO: Add any initialization after the InitializeComponent call
            ttpMailPage.SetToolTip( lnkFrom,  "Load address from file - multiple addresses in one field");
            ttpMailPage.SetToolTip( lnkTo,    "Load address from file - multiple addresses in one field");
            ttpMailPage.SetToolTip( lnkCC,    "Load address from file - multiple addresses in one field");
            ttpMailPage.SetToolTip( lnkBCC,   "Load address from file - multiple addresses in one field");
            ttpMailPage.SetToolTip( lnkAttach,"Load address from file - multiple addresses in one field");

            commObj.InitComboBoxItem( cboFrom, "[From Address]" );
            commObj.InitComboBoxItem( cboTo, "[To Address]" );
            commObj.InitComboBoxItem( cboCC, "[CC Address]" );
            commObj.InitComboBoxItem( cboBCC, "[BCC Address]" );
            commObj.InitComboBoxItem( cboSMTP, "[SMTP IP]" );
            commObj.InitComboBoxItem( cboPort, "[Port]" );

            commObj.InitComboBoxItem( cboHeader1, "[Z Header 1]" );
            commObj.InitComboBoxItem( cboHeaderVal1, "[Z Value 1]" );
            commObj.InitComboBoxItem( cboHeader2, "[Z Header 2]" );
            commObj.InitComboBoxItem( cboHeaderVal2, "[Z Value 2]" );
        }// end of constructor

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            if( disposing )
            {
                Debug.WriteLine("MailPage.cs - Deposing MailPage Object");
                commObj.LogToFile("MailPage.cs - Deposing MailPage Object");
                if(components != null)
                {
                    Debug.WriteLine("\t Dispose component");
                    commObj.LogToFile("\t Dispose component");
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
            this.lnkFrom = new System.Windows.Forms.LinkLabel();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.lnkAttach = new System.Windows.Forms.LinkLabel();
            this.txtAttach = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.lblSMTP = new System.Windows.Forms.Label();
            this.lblPort = new System.Windows.Forms.Label();
            this.cboSMTP = new System.Windows.Forms.ComboBox();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.btnTest = new System.Windows.Forms.Button();
            this.ttpMailPage = new System.Windows.Forms.ToolTip(this.components);
            this.chkConSearch = new System.Windows.Forms.CheckBox();
            this.cboHeader1 = new System.Windows.Forms.ComboBox();
            this.cboHeader2 = new System.Windows.Forms.ComboBox();
            this.cboHeaderVal1 = new System.Windows.Forms.ComboBox();
            this.cboHeaderVal2 = new System.Windows.Forms.ComboBox();
            this.chkMultiAttach = new System.Windows.Forms.CheckBox();
            this.rdoMailAPI = new System.Windows.Forms.RadioButton();
            this.rdoSMTPClient = new System.Windows.Forms.RadioButton();
            this.dtpSendDate = new System.Windows.Forms.DateTimePicker();
            this.chkUnindexable = new System.Windows.Forms.CheckBox();
            this.chkHeader = new System.Windows.Forms.CheckBox();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.cboFrom = new System.Windows.Forms.ComboBox();
            this.chkUnparsable = new System.Windows.Forms.CheckBox();
            this.lblHeader = new System.Windows.Forms.Label();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblCustSendDate = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lnkFrom
            // 
            this.lnkFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFrom.Location = new System.Drawing.Point(4, 8);
            this.lnkFrom.Name = "lnkFrom";
            this.lnkFrom.Size = new System.Drawing.Size(56, 16);
            this.lnkFrom.TabIndex = 0;
            this.lnkFrom.TabStop = true;
            this.lnkFrom.Text = "From :";
            this.lnkFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkFrom.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFrom_LinkClicked);
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(4, 32);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(56, 16);
            this.lnkTo.TabIndex = 1;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkTo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkTo_LinkClicked);
            // 
            // lnkCC
            // 
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(4, 56);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(56, 16);
            this.lnkCC.TabIndex = 2;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkCC.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCC_LinkClicked);
            // 
            // lnkBCC
            // 
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(188, 56);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(40, 16);
            this.lnkBCC.TabIndex = 3;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkBCC.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkBCC_LinkClicked);
            // 
            // txtSubject
            // 
            this.txtSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.txtSubject.Location = new System.Drawing.Point(68, 76);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(312, 20);
            this.txtSubject.TabIndex = 9;
            this.txtSubject.Text = "txtSubject";
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(4, 80);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 10;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(8, 324);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(372, 100);
            this.richBox.TabIndex = 11;
            this.richBox.Text = "richBox";
            // 
            // lnkAttach
            // 
            this.lnkAttach.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkAttach.Location = new System.Drawing.Point(64, 92);
            this.lnkAttach.Name = "lnkAttach";
            this.lnkAttach.Size = new System.Drawing.Size(68, 16);
            this.lnkAttach.TabIndex = 12;
            this.lnkAttach.TabStop = true;
            this.lnkAttach.Text = "Attachment : ";
            this.lnkAttach.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpMailPage.SetToolTip(this.lnkAttach, "Select ONE attachment");
            this.lnkAttach.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkAttach_LinkClicked);
            // 
            // txtAttach
            // 
            this.txtAttach.Location = new System.Drawing.Point(136, 88);
            this.txtAttach.Name = "txtAttach";
            this.txtAttach.Size = new System.Drawing.Size(240, 21);
            this.txtAttach.TabIndex = 13;
            this.txtAttach.Text = "";
            this.ttpMailPage.SetToolTip(this.txtAttach, "Full path of attachment file. Separated by semicolon");
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(324, 268);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(56, 21);
            this.btnSend.TabIndex = 14;
            this.btnSend.Text = "Send";
            this.ttpMailPage.SetToolTip(this.btnSend, "Initiate send thread");
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // lblSMTP
            // 
            this.lblSMTP.Location = new System.Drawing.Point(8, 300);
            this.lblSMTP.Name = "lblSMTP";
            this.lblSMTP.Size = new System.Drawing.Size(80, 16);
            this.lblSMTP.TabIndex = 15;
            this.lblSMTP.Text = "SMTP Server";
            this.ttpMailPage.SetToolTip(this.lblSMTP, "SMTP server name or IP address");
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(232, 300);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(38, 16);
            this.lblPort.TabIndex = 16;
            this.lblPort.Text = "Port # ";
            this.ttpMailPage.SetToolTip(this.lblPort, "SMTP port number");
            // 
            // cboSMTP
            // 
            this.cboSMTP.DisplayMember = "10.1.11.15";
            this.cboSMTP.ItemHeight = 15;
            this.cboSMTP.Items.AddRange(new object[] {
                                                         ""});
            this.cboSMTP.Location = new System.Drawing.Point(92, 296);
            this.cboSMTP.Name = "cboSMTP";
            this.cboSMTP.Size = new System.Drawing.Size(132, 23);
            this.cboSMTP.Sorted = true;
            this.cboSMTP.TabIndex = 17;
            this.cboSMTP.Text = "10.1.89.201";
            this.ttpMailPage.SetToolTip(this.cboSMTP, "Server name or IP address");
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 15;
            this.cboPort.Location = new System.Drawing.Point(276, 296);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(44, 23);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 18;
            this.cboPort.Text = "25";
            this.ttpMailPage.SetToolTip(this.cboPort, "SMTP port number");
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(324, 296);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(56, 21);
            this.btnTest.TabIndex = 19;
            this.btnTest.Text = "Test";
            this.ttpMailPage.SetToolTip(this.btnTest, "Test the SMTP connection");
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // chkConSearch
            // 
            this.chkConSearch.Enabled = false;
            this.chkConSearch.Location = new System.Drawing.Point(256, 124);
            this.chkConSearch.Name = "chkConSearch";
            this.chkConSearch.Size = new System.Drawing.Size(116, 16);
            this.chkConSearch.TabIndex = 21;
            this.chkConSearch.Text = "Content Search";
            this.chkConSearch.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.ttpMailPage.SetToolTip(this.chkConSearch, "Able to search by content");
            // 
            // cboHeader1
            // 
            this.cboHeader1.Enabled = false;
            this.cboHeader1.Items.AddRange(new object[] {
                                                            "X-ZANTAZ-SS-NUM",
                                                            "X-ZANTAZ-SYMBOL",
                                                            "X-ZANTAZ-USER"});
            this.cboHeader1.Location = new System.Drawing.Point(68, 140);
            this.cboHeader1.Name = "cboHeader1";
            this.cboHeader1.Size = new System.Drawing.Size(180, 23);
            this.cboHeader1.TabIndex = 23;
            this.cboHeader1.Text = "Content-Identifer";
            this.ttpMailPage.SetToolTip(this.cboHeader1, "Just an example");
            // 
            // cboHeader2
            // 
            this.cboHeader2.Enabled = false;
            this.cboHeader2.Items.AddRange(new object[] {
                                                            "X-ZANTAZ-SS-NUM",
                                                            "X-ZANTAZ-SYMBOL",
                                                            "X-ZANTAZ-USER"});
            this.cboHeader2.Location = new System.Drawing.Point(68, 164);
            this.cboHeader2.Name = "cboHeader2";
            this.cboHeader2.Size = new System.Drawing.Size(180, 23);
            this.cboHeader2.TabIndex = 24;
            this.cboHeader2.Text = "X-ZANTAZ-ACCOUNT-CODE";
            this.ttpMailPage.SetToolTip(this.cboHeader2, "Just an example");
            // 
            // cboHeaderVal1
            // 
            this.cboHeaderVal1.Enabled = false;
            this.cboHeaderVal1.Items.AddRange(new object[] {
                                                               "123-456-1234",
                                                               "ABC",
                                                               "User Name",
                                                               "ExJournalReport"});
            this.cboHeaderVal1.Location = new System.Drawing.Point(252, 140);
            this.cboHeaderVal1.Name = "cboHeaderVal1";
            this.cboHeaderVal1.Size = new System.Drawing.Size(128, 23);
            this.cboHeaderVal1.TabIndex = 25;
            this.cboHeaderVal1.Text = "ExJournalReport";
            this.ttpMailPage.SetToolTip(this.cboHeaderVal1, "Just an example");
            // 
            // cboHeaderVal2
            // 
            this.cboHeaderVal2.Enabled = false;
            this.cboHeaderVal2.Items.AddRange(new object[] {
                                                               "123-456-1234",
                                                               "ABC",
                                                               "User Name"});
            this.cboHeaderVal2.Location = new System.Drawing.Point(252, 164);
            this.cboHeaderVal2.Name = "cboHeaderVal2";
            this.cboHeaderVal2.Size = new System.Drawing.Size(128, 23);
            this.cboHeaderVal2.TabIndex = 26;
            this.cboHeaderVal2.Text = "Z123";
            this.ttpMailPage.SetToolTip(this.cboHeaderVal2, "Example too");
            // 
            // chkMultiAttach
            // 
            this.chkMultiAttach.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.chkMultiAttach.Location = new System.Drawing.Point(52, 88);
            this.chkMultiAttach.Name = "chkMultiAttach";
            this.chkMultiAttach.Size = new System.Drawing.Size(16, 20);
            this.chkMultiAttach.TabIndex = 31;
            this.chkMultiAttach.Text = "Multiple";
            this.chkMultiAttach.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.ttpMailPage.SetToolTip(this.chkMultiAttach, "Set multple attachments");
            // 
            // rdoMailAPI
            // 
            this.rdoMailAPI.Checked = true;
            this.rdoMailAPI.Location = new System.Drawing.Point(12, 104);
            this.rdoMailAPI.Name = "rdoMailAPI";
            this.rdoMailAPI.Size = new System.Drawing.Size(92, 16);
            this.rdoMailAPI.TabIndex = 32;
            this.rdoMailAPI.TabStop = true;
            this.rdoMailAPI.Text = "MS Mail API";
            this.ttpMailPage.SetToolTip(this.rdoMailAPI, "Use Microsoft Mail API to send mail");
            this.rdoMailAPI.CheckedChanged += new System.EventHandler(this.rdoMailAPI_CheckedChanged);
            // 
            // rdoSMTPClient
            // 
            this.rdoSMTPClient.Location = new System.Drawing.Point(12, 224);
            this.rdoSMTPClient.Name = "rdoSMTPClient";
            this.rdoSMTPClient.Size = new System.Drawing.Size(92, 16);
            this.rdoSMTPClient.TabIndex = 33;
            this.rdoSMTPClient.Text = "SMTP Client";
            this.ttpMailPage.SetToolTip(this.rdoSMTPClient, "Use SMTP to send simple mail");
            this.rdoSMTPClient.CheckedChanged += new System.EventHandler(this.rdoSMTPClient_CheckedChanged);
            // 
            // dtpSendDate
            // 
            this.dtpSendDate.CustomFormat = "MM/dd/yyyy hh:mm tt";
            this.dtpSendDate.Enabled = false;
            this.dtpSendDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpSendDate.Location = new System.Drawing.Point(148, 12);
            this.dtpSendDate.Name = "dtpSendDate";
            this.dtpSendDate.Size = new System.Drawing.Size(156, 21);
            this.dtpSendDate.TabIndex = 55;
            this.ttpMailPage.SetToolTip(this.dtpSendDate, "SMTP only - Send date");
            // 
            // chkUnindexable
            // 
            this.chkUnindexable.Enabled = false;
            this.chkUnindexable.Location = new System.Drawing.Point(208, 40);
            this.chkUnindexable.Name = "chkUnindexable";
            this.chkUnindexable.Size = new System.Drawing.Size(96, 20);
            this.chkUnindexable.TabIndex = 59;
            this.chkUnindexable.Text = "Unindexable";
            this.chkUnindexable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.ttpMailPage.SetToolTip(this.chkUnindexable, "** Disabled **");
            // 
            // chkHeader
            // 
            this.chkHeader.Location = new System.Drawing.Point(68, 124);
            this.chkHeader.Name = "chkHeader";
            this.chkHeader.Size = new System.Drawing.Size(132, 16);
            this.chkHeader.TabIndex = 20;
            this.chkHeader.Text = "Set Custom Header";
            this.chkHeader.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.ttpMailPage.SetToolTip(this.chkHeader, "Set custom info in mail header");
            this.chkHeader.CheckedChanged += new System.EventHandler(this.chkHeader_CheckedChanged);
            // 
            // cboTo
            // 
            this.cboTo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
                | System.Windows.Forms.AnchorStyles.Right)));
            this.cboTo.Location = new System.Drawing.Point(68, 25);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(312, 23);
            this.cboTo.TabIndex = 27;
            this.cboTo.Text = "login0@company1.zantaz.com";
            this.ttpMailPage.SetToolTip(this.cboTo, "Rcpt To: select from the pull down box or load multiple from the link");
            // 
            // cboFrom
            // 
            this.cboFrom.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.cboFrom.Location = new System.Drawing.Point(68, 1);
            this.cboFrom.Name = "cboFrom";
            this.cboFrom.Size = new System.Drawing.Size(312, 23);
            this.cboFrom.TabIndex = 28;
            this.cboFrom.Text = "kent@zel.com";
            this.ttpMailPage.SetToolTip(this.cboFrom, "Mail From - select from the pull down box");
            // 
            // chkUnparsable
            // 
            this.chkUnparsable.Enabled = false;
            this.chkUnparsable.Location = new System.Drawing.Point(80, 40);
            this.chkUnparsable.Name = "chkUnparsable";
            this.chkUnparsable.Size = new System.Drawing.Size(96, 20);
            this.chkUnparsable.TabIndex = 58;
            this.chkUnparsable.Text = "Unparseable";
            this.chkUnparsable.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.ttpMailPage.SetToolTip(this.chkUnparsable, "Am I retired?");
            // 
            // lblHeader
            // 
            this.lblHeader.Enabled = false;
            this.lblHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblHeader.Location = new System.Drawing.Point(8, 140);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Size = new System.Drawing.Size(56, 16);
            this.lblHeader.TabIndex = 22;
            this.lblHeader.Text = "Header :";
            this.lblHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cboCC
            // 
            this.cboCC.Location = new System.Drawing.Point(68, 49);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(124, 23);
            this.cboCC.TabIndex = 29;
            // 
            // cboBCC
            // 
            this.cboBCC.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
                | System.Windows.Forms.AnchorStyles.Right)));
            this.cboBCC.Location = new System.Drawing.Point(228, 49);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(152, 23);
            this.cboBCC.TabIndex = 30;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkMultiAttach);
            this.groupBox1.Controls.Add(this.lnkAttach);
            this.groupBox1.Controls.Add(this.txtAttach);
            this.groupBox1.Location = new System.Drawing.Point(4, 104);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(380, 116);
            this.groupBox1.TabIndex = 56;
            this.groupBox1.TabStop = false;
            // 
            // lblCustSendDate
            // 
            this.lblCustSendDate.Enabled = false;
            this.lblCustSendDate.Location = new System.Drawing.Point(80, 16);
            this.lblCustSendDate.Name = "lblCustSendDate";
            this.lblCustSendDate.Size = new System.Drawing.Size(64, 16);
            this.lblCustSendDate.TabIndex = 57;
            this.lblCustSendDate.Text = "Send Date";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkUnindexable);
            this.groupBox2.Controls.Add(this.chkUnparsable);
            this.groupBox2.Controls.Add(this.lblCustSendDate);
            this.groupBox2.Controls.Add(this.dtpSendDate);
            this.groupBox2.Location = new System.Drawing.Point(4, 224);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(312, 68);
            this.groupBox2.TabIndex = 58;
            this.groupBox2.TabStop = false;
            // 
            // MailPage
            // 
            this.Controls.Add(this.rdoSMTPClient);
            this.Controls.Add(this.rdoMailAPI);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.cboFrom);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.cboHeaderVal2);
            this.Controls.Add(this.cboHeaderVal1);
            this.Controls.Add(this.cboHeader2);
            this.Controls.Add(this.cboHeader1);
            this.Controls.Add(this.lblHeader);
            this.Controls.Add(this.chkConSearch);
            this.Controls.Add(this.chkHeader);
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
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnSend);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Name = "MailPage";
            this.Size = new System.Drawing.Size(388, 428);
            this.Load += new System.EventHandler(this.MailPage_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        private void MailPage_Load(object sender, System.EventArgs e)
        {
		
        }

        private void lnkFrom_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - lnkFrom_LinkClicked" );

            string str = "";
            commObj.LoadAddrFromFile( ref str );
            cboFrom.Text = str;			        		
        }// end of lnkFrom_LinkClicked

        private void lnkTo_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - lnkTo_LinkClicked" );

            string str = ""; // temp storage
            commObj.LoadAddrFromFile( ref str ); // pass by ref
            cboTo.Text = str; // assign back to text box        
		
        }//end of lnkTo_LinkClicked

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

        private void lnkBCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - lnkBCC_LinkClicked" );

            string str = "";
            commObj.LoadAddrFromFile( ref str );
            cboBCC.Text = str;			        		
        }//end of lnkBCC_LinkClicked

        private void lnkCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - lnkCC_LinkClicked" );

            string str = "";
            commObj.LoadAddrFromFile( ref str );
            cboCC.Text = str;        		
        }//end of lnkCC_LinkClicked

        private void btnTest_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - lnkTest_LinkClicked" );
            this.Cursor = Cursors.WaitCursor;

            richBox.Text = commObj.TestSMTPConnection(cboSMTP.Text, cboPort.Text)?"Connection OK":"Connection FAIL";

            this.Cursor = Cursors.Default;
        }//end of btnTest_Click

        private void btnSend_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "MailPage.cs - btnSend_Click" );
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;

            if( rdoSMTPClient.Checked )
            {
                SMTPClientSender();
            }//end of if - do the SMTP Client Sending
            else
                if( rdoMailAPI.Checked )
				{
					MailAPISender();
				}// end of if - use the MS Mail api sending

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
            //            richBox.Text = "Send mail Done!";
        }//end of btnSend_Click

        private void chkHeader_CheckedChanged(object sender, System.EventArgs e)
        {
            if( chkHeader.Checked )
            {
                lblHeader.Enabled     = true;
                chkConSearch.Enabled  = true;
                cboHeader1.Enabled    = true;
                cboHeader2.Enabled    = true;
                cboHeaderVal1.Enabled = true;
                cboHeaderVal2.Enabled = true;
            }//end of if header checked
            else
            {
                lblHeader.Enabled     = false;
                chkConSearch.Enabled  = false;
                cboHeader1.Enabled    = false;
                cboHeader2.Enabled    = false;
                cboHeaderVal1.Enabled = false;
                cboHeaderVal2.Enabled = false;
            }//end of else
        }

        /// <summary>
        /// User MS Mail API to send mail - it is default method to send mail
        /// </summary>
        public void MailAPISender()
        {
            Trace.WriteLine( "MailPage.cs - MailAPISender" );
			string savBody = richBox.Text; // save the content of the rich Box - custom info
            try
            {
                MailMessage mailMsg = new MailMessage();
                mailMsg.From    = cboFrom.Text;
                mailMsg.To      = cboTo.Text;
                mailMsg.Cc      = cboCC.Text;
                mailMsg.Bcc     = cboBCC.Text;
                mailMsg.Subject = txtSubject.Text;
                mailMsg.Body    = savBody		// this will not change 
                    + "\nFrom: " + cboFrom.Text // start from here change
                    + "\n TO:  " + cboTo.Text
                    + "\n CC:  " + cboCC.Text
                    + "\n BCC: " + cboBCC.Text
                    + "\n Body_Subject: " + txtSubject.Text
                    + "\n" + DateTime.Now;

                if( chkHeader.Checked )
                {
                    // save for reference - this one is needed for creating new doc type
                    //                    mailMsg.Headers.Add("X-ZANTAZDOCCLASS", "CONFIRMATIONS");
                    mailMsg.Headers.Add( cboHeader1.Text, cboHeaderVal1.Text );
                    mailMsg.Headers.Add( cboHeader2.Text, cboHeaderVal2.Text );
                    //kentest
                    //                    mailMsg.Headers.Add( "Recipient:", "faker@faker.com" );

                    if( chkConSearch.Checked )
                        mailMsg.Headers.Add("=BODY=", "");
                }//end of if
				
                // validate the input - trim the space before check the length
                txtAttach.Text.TrimStart( new char[] {' '} );
                if( 0 < txtAttach.Text.Length )
                {
                    if( chkMultiAttach.Checked )
                    {
                        char[] delim = new char[]{';'};
                        foreach( string str in txtAttach.Text.Split(delim) )
                        {
                            mailMsg.Attachments.Add( new MailAttachment(str, MailEncoding.Base64) );
                        }//end of foreach
                    }// end of if - multiple attachment
                    else
                    {
                        mailMsg.Attachments.Add( new MailAttachment(txtAttach.Text, MailEncoding.Base64) );
                    }//end of else
                }//end of if - attachment

                // If SmtpServer is not set, local SMTP server is used
                SmtpMail.SmtpServer = cboSMTP.Text;

                // Test the connection - if smtp server down, exit.
                System.Net.Sockets.TcpClient tcpClient = new TcpClient();
                tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboPort.Text) );

                SmtpMail.Send( mailMsg );
                MessageBox.Show("Message Sent to " + cboTo.Text, msgCaption );

                tcpClient.Close();

            }// end of try
            catch( System.Web.HttpException ex )
            {
                Trace.WriteLine( "\tMailAPISender() - HTTP Exception: " + ex.Message.ToString() );
                MessageBox.Show( ex.Message.ToString(), msgCaption);
            }//end of catch - HttpException
            catch( Exception gex )
            {
                Trace.WriteLine( "\tMailPage.cs - Generic Exception: " + gex.Message.ToString() );
                MessageBox.Show( gex.Message.ToString(), msgCaption);
            }//end of catch - generic exception
        }// end of MailAPISender

        /// <summary>
        /// Use Custom SMTP client - SMTPSender to send mail.... For de-duplication testing.
        /// It is special for audit center to send two identical mail to different repositories.
        /// No attachment, no zantaz custom header... Just a simple mail program.
        /// </summary>
        public void SMTPClientSender()
        {
            Trace.WriteLine("MailPage.cs - SMTPClientSender()");
            try
            {
                QATool.SMTPSender smtpSender = new SMTPSender();
                smtpSender.mailFrom    = cboFrom.Text;
                smtpSender.mailTo      = cboTo.Text;
                smtpSender.mailCC      = cboCC.Text;
                smtpSender.mailBCC     = cboBCC.Text;
                smtpSender.mailSentDate= dtpSendDate.Text;
                smtpSender.mailSubject = txtSubject.Text;
                smtpSender.mailBody    = richBox.Text;

                smtpSender.smtpServer  = cboSMTP.Text;
                smtpSender.smtpPortNum = cboPort.Text;

                if( chkUnindexable.Checked )
                {
//save for refer    smtpSender.mailBody = RandString();
                }

                if( chkUnparsable.Checked )
                {
                    smtpSender.mailContentType = "unknowFoo";
                }

                smtpSender.SmtpSend();
            }//end of try
            catch( QATool.SMTPException ex )
            {
                Trace.WriteLine("\tSMTPClientSender() Exception: " + ex.SmtpMessage.ToString() );
                MessageBox.Show( ex.SmtpMessage.ToString(), msgCaption );
            }//end of catch
        }//end of SMTPClientSender

        /// <summary>
        /// Disable some of the controls
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rdoSMTPClient_CheckedChanged(object sender, System.EventArgs e)
        {
            chkHeader.Enabled     = false; // special case for this check box

            lblHeader.Enabled     = false;
            chkConSearch.Enabled  = false;
            cboHeader1.Enabled    = false;
            cboHeader2.Enabled    = false;
            cboHeaderVal1.Enabled = false;
            cboHeaderVal2.Enabled = false;

            txtAttach.Enabled     = false;
            lnkAttach.Enabled     = false;
            chkMultiAttach.Enabled= false;

            dtpSendDate.Enabled     = true; // send date for SMTP client mail ONLY
            lblCustSendDate.Enabled = true;
            chkUnparsable.Enabled   = true;
            chkUnindexable.Enabled  = true;
        }//end of rdoSMTPClient_CheckedChanged

        /// <summary>
        /// Enable some of the controls
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rdoMailAPI_CheckedChanged(object sender, System.EventArgs e)
        {
            chkHeader.Enabled     = true; // special case for this check box

            if( chkHeader.Checked )
            {
                lblHeader.Enabled     = true;
                chkConSearch.Enabled  = true;
                cboHeader1.Enabled    = true;
                cboHeader2.Enabled    = true;
                cboHeaderVal1.Enabled = true;
                cboHeaderVal2.Enabled = true;
            }//end of if 

            txtAttach.Enabled      = true;
            lnkAttach.Enabled      = true;
            chkMultiAttach.Enabled = true;

            dtpSendDate.Enabled    = false; // NOT avaliable to MS Mail API
            lblCustSendDate.Enabled= false;
            chkUnparsable.Enabled  = false;
            chkUnindexable.Enabled = false;

        }// end of rdoMailAPI_CheckedChanged

/*** save for reference
        // Create a random object with a timer-generated seed.
        public string RandString()
        {
            // Wait to allow the timer to advance.
            Thread.Sleep( 1 );
            Random autoRand = new Random( );

            return( GenIntRandNum( autoRand ) );
        }

        // Generate random numbers from the specified Random object.
        public string GenIntRandNum( Random randObj )
        {
            string str = "";
            string longstr = "";
            // Generate the first six random integers.
            for( int j = 0; j < 500000; j++ )
            {
                str = j.ToString() + ": " + randObj.Next().ToString();
                commObj.LogGUID( "Token.txt", str );
            }

            return( str );
        }
** save for reference **/        



	}
}
