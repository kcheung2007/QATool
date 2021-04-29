using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Mail;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for ZelMsgPage.
	/// </summary>
	public class ZelMsgPage : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.RichTextBox richBox;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.ComboBox cboSMTP;
        private System.Windows.Forms.PropertyGrid propGrid1;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.ToolTip ttpZELTool;
        private System.Windows.Forms.TextBox txtListFile;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.LinkLabel lnkListFile;
        private System.Windows.Forms.LinkLabel lnkFolder;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.Label lblSMTP;
        private System.Windows.Forms.ComboBox cboCC;
        private System.Windows.Forms.ComboBox cboBCC;
        private System.Windows.Forms.ComboBox cboTo;
        private System.Windows.Forms.LinkLabel lnkBCC;
        private System.Windows.Forms.LinkLabel lnkCC;
        private System.Windows.Forms.LinkLabel lnkTo;
        private System.Windows.Forms.LinkLabel lnkFrom;
        private System.Windows.Forms.GroupBox gbxDigiSafe;
        private System.Windows.Forms.CheckBox chkGUID;
        private System.Windows.Forms.ComboBox cboFrom;
        private System.Windows.Forms.Button btnAbort;
        private System.ComponentModel.IContainer components;

        private string m_dataFolder = "";
        private string m_listFile   = "";

        private QATool.CommObj       commObj = new CommObj();
        private QATool.ZelVarDataObj varObj  = new ZelVarDataObj();
        private System.Windows.Forms.CheckBox chkModify;
        private System.Windows.Forms.RadioButton rdoMSAPI;
        private System.Windows.Forms.RadioButton rdoSMTPClient;
        private Thread sendZelMailThread;
        private Thread sendSmtpMailThread;
        private Thread sendCustMailThread;

        const int BYTE_SIZE = 8192;


		public ZelMsgPage()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            m_listFile = txtListFile.Text;
            m_dataFolder = txtFolder.Text;
            Directory.SetCurrentDirectory( Application.StartupPath );

            commObj.InitComboBoxItem( cboFrom, "[From Address]" );
            commObj.InitComboBoxItem( cboTo,   "[To Address]" );
            commObj.InitComboBoxItem( cboCC,   "[CC Address]" );
            commObj.InitComboBoxItem( cboBCC,  "[BCC Address]" );
            commObj.InitComboBoxItem( cboSMTP, "[SMTP IP]" );
            commObj.InitComboBoxItem( cboPort, "[Port]" );
		}// end of constructor

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
            Debug.WriteLine( "QATool.cs - Depose Object");
            commObj.LogToFile("QATool.cs - Depose Object");

            if( sendZelMailThread != null && sendZelMailThread.IsAlive )
                this.KillSendZelMailThread();

            if( sendSmtpMailThread != null && sendSmtpMailThread.IsAlive )
                this.KillSendSmtpMailThread();

            if( sendCustMailThread != null && sendCustMailThread.IsAlive )
                this.KillSendCustMailThread();

            if( disposing )
            {
                if(components != null)
                {                    					
                    components.Dispose();
                    commObj.LogToFile("\tDispose ZelMsgPage component");
                    Debug.WriteLine("\t Dispose ZelMsgPage component");
                }
            }// end of if
            base.Dispose( disposing );
		}// end of Dispose

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.lblSubject = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.richBox = new System.Windows.Forms.RichTextBox();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.cboSMTP = new System.Windows.Forms.ComboBox();
            this.propGrid1 = new System.Windows.Forms.PropertyGrid();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.ttpZELTool = new System.Windows.Forms.ToolTip(this.components);
            this.txtListFile = new System.Windows.Forms.TextBox();
            this.btnTest = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.lnkListFile = new System.Windows.Forms.LinkLabel();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.chkGUID = new System.Windows.Forms.CheckBox();
            this.btnAbort = new System.Windows.Forms.Button();
            this.rdoMSAPI = new System.Windows.Forms.RadioButton();
            this.rdoSMTPClient = new System.Windows.Forms.RadioButton();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblSMTP = new System.Windows.Forms.Label();
            this.cboCC = new System.Windows.Forms.ComboBox();
            this.cboBCC = new System.Windows.Forms.ComboBox();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.lnkBCC = new System.Windows.Forms.LinkLabel();
            this.lnkCC = new System.Windows.Forms.LinkLabel();
            this.lnkTo = new System.Windows.Forms.LinkLabel();
            this.lnkFrom = new System.Windows.Forms.LinkLabel();
            this.gbxDigiSafe = new System.Windows.Forms.GroupBox();
            this.cboFrom = new System.Windows.Forms.ComboBox();
            this.chkModify = new System.Windows.Forms.CheckBox();
            this.gbxDigiSafe.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblSubject
            // 
            this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lblSubject.Location = new System.Drawing.Point(10, 124);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(56, 16);
            this.lblSubject.TabIndex = 144;
            this.lblSubject.Text = "Subject :";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtSubject
            // 
            this.txtSubject.Enabled = false;
            this.txtSubject.Location = new System.Drawing.Point(70, 120);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(256, 20);
            this.txtSubject.TabIndex = 143;
            this.txtSubject.Text = "txtSubject";
            // 
            // richBox
            // 
            this.richBox.Location = new System.Drawing.Point(2, 360);
            this.richBox.Name = "richBox";
            this.richBox.Size = new System.Drawing.Size(292, 64);
            this.richBox.TabIndex = 142;
            this.richBox.Text = "richBox";
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 13;
            this.cboPort.Location = new System.Drawing.Point(242, 332);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(52, 21);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 139;
            this.cboPort.Text = "25";
            // 
            // cboSMTP
            // 
            this.cboSMTP.ItemHeight = 13;
            this.cboSMTP.Items.AddRange(new object[] {
                                                         "10.1.11.40",
                                                         "10.1.41.234",
                                                         "10.1.42.100",
                                                         "10.1.42.102"});
            this.cboSMTP.Location = new System.Drawing.Point(66, 332);
            this.cboSMTP.Name = "cboSMTP";
            this.cboSMTP.Size = new System.Drawing.Size(144, 21);
            this.cboSMTP.Sorted = true;
            this.cboSMTP.TabIndex = 138;
            this.cboSMTP.Text = "10.1.42.102";
            this.ttpZELTool.SetToolTip(this.cboSMTP, "SMTP server name or IP");
            // 
            // propGrid1
            // 
            this.propGrid1.CommandsVisibleIfAvailable = true;
            this.propGrid1.HelpVisible = false;
            this.propGrid1.LargeButtons = false;
            this.propGrid1.LineColor = System.Drawing.SystemColors.ScrollBar;
            this.propGrid1.Location = new System.Drawing.Point(6, 228);
            this.propGrid1.Name = "propGrid1";
            this.propGrid1.PropertySort = System.Windows.Forms.PropertySort.Alphabetical;
            this.propGrid1.Size = new System.Drawing.Size(376, 52);
            this.propGrid1.TabIndex = 146;
            this.propGrid1.Text = "propertyGrid1";
            this.propGrid1.ToolbarVisible = false;
            this.propGrid1.ViewBackColor = System.Drawing.SystemColors.Window;
            this.propGrid1.ViewForeColor = System.Drawing.SystemColors.WindowText;
            // 
            // txtFolder
            // 
            this.txtFolder.Location = new System.Drawing.Point(122, 308);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(208, 20);
            this.txtFolder.TabIndex = 137;
            this.txtFolder.Text = "D:\\TestData\\List";
            this.ttpZELTool.SetToolTip(this.txtFolder, "Messages in text file format");
            // 
            // txtListFile
            // 
            this.txtListFile.Location = new System.Drawing.Point(74, 284);
            this.txtListFile.Name = "txtListFile";
            this.txtListFile.Size = new System.Drawing.Size(256, 20);
            this.txtListFile.TabIndex = 145;
            this.txtListFile.Text = "D:\\TestData\\List\\aspecial.txt";
            this.ttpZELTool.SetToolTip(this.txtListFile, "Text File contains a list of message file name");
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(298, 332);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(80, 20);
            this.btnTest.TabIndex = 141;
            this.btnTest.Text = "Test";
            this.ttpZELTool.SetToolTip(this.btnTest, "Test the SMTP connection");
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(298, 380);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(80, 20);
            this.btnSend.TabIndex = 133;
            this.btnSend.Text = "Send";
            this.ttpZELTool.SetToolTip(this.btnSend, "By MS API unless above SMTP was selected");
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // lnkListFile
            // 
            this.lnkListFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkListFile.Location = new System.Drawing.Point(14, 288);
            this.lnkListFile.Name = "lnkListFile";
            this.lnkListFile.Size = new System.Drawing.Size(52, 16);
            this.lnkListFile.TabIndex = 140;
            this.lnkListFile.TabStop = true;
            this.lnkListFile.Text = "List File:";
            this.lnkListFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpZELTool.SetToolTip(this.lnkListFile, "Browse the location of list file");
            this.lnkListFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkListFile_LinkClicked);
            // 
            // lnkFolder
            // 
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.Location = new System.Drawing.Point(66, 312);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(44, 16);
            this.lnkFolder.TabIndex = 136;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder:";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpZELTool.SetToolTip(this.lnkFolder, "Point to the message files folder");
            this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
            // 
            // chkGUID
            // 
            this.chkGUID.Enabled = false;
            this.chkGUID.Location = new System.Drawing.Point(324, 120);
            this.chkGUID.Name = "chkGUID";
            this.chkGUID.Size = new System.Drawing.Size(52, 16);
            this.chkGUID.TabIndex = 122;
            this.chkGUID.Text = "GUID";
            this.ttpZELTool.SetToolTip(this.chkGUID, "Include GUID in the subject");
            // 
            // btnAbort
            // 
            this.btnAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAbort.Location = new System.Drawing.Point(298, 404);
            this.btnAbort.Name = "btnAbort";
            this.btnAbort.Size = new System.Drawing.Size(80, 20);
            this.btnAbort.TabIndex = 124;
            this.btnAbort.Text = "Abort";
            this.ttpZELTool.SetToolTip(this.btnAbort, "Kill the sending thread");
            this.btnAbort.Click += new System.EventHandler(this.btnAbort_Click);
            // 
            // rdoMSAPI
            // 
            this.rdoMSAPI.Location = new System.Drawing.Point(120, 208);
            this.rdoMSAPI.Name = "rdoMSAPI";
            this.rdoMSAPI.Size = new System.Drawing.Size(84, 16);
            this.rdoMSAPI.TabIndex = 123;
            this.rdoMSAPI.Text = "Use MS API";
            this.ttpZELTool.SetToolTip(this.rdoMSAPI, "As message body");
            this.rdoMSAPI.Click += new System.EventHandler(this.rdoMSAPI_Click);
            // 
            // rdoSMTPClient
            // 
            this.rdoSMTPClient.Checked = true;
            this.rdoSMTPClient.Location = new System.Drawing.Point(216, 208);
            this.rdoSMTPClient.Name = "rdoSMTPClient";
            this.rdoSMTPClient.Size = new System.Drawing.Size(116, 16);
            this.rdoSMTPClient.TabIndex = 148;
            this.rdoSMTPClient.TabStop = true;
            this.rdoSMTPClient.Text = "Use SMTP Client";
            this.ttpZELTool.SetToolTip(this.rdoSMTPClient, "Stream into socket");
            this.rdoSMTPClient.Click += new System.EventHandler(this.rdoSMTPClient_Click);
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(210, 336);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(28, 16);
            this.lblPort.TabIndex = 135;
            this.lblPort.Text = "Port";
            // 
            // lblSMTP
            // 
            this.lblSMTP.Location = new System.Drawing.Point(22, 336);
            this.lblSMTP.Name = "lblSMTP";
            this.lblSMTP.Size = new System.Drawing.Size(36, 16);
            this.lblSMTP.TabIndex = 134;
            this.lblSMTP.Text = "SMTP";
            // 
            // cboCC
            // 
            this.cboCC.Enabled = false;
            this.cboCC.Location = new System.Drawing.Point(70, 72);
            this.cboCC.Name = "cboCC";
            this.cboCC.Size = new System.Drawing.Size(312, 21);
            this.cboCC.TabIndex = 130;
            // 
            // cboBCC
            // 
            this.cboBCC.Enabled = false;
            this.cboBCC.Location = new System.Drawing.Point(70, 96);
            this.cboBCC.Name = "cboBCC";
            this.cboBCC.Size = new System.Drawing.Size(312, 21);
            this.cboBCC.TabIndex = 131;
            // 
            // cboTo
            // 
            this.cboTo.Items.AddRange(new object[] {
                                                       ""});
            this.cboTo.Location = new System.Drawing.Point(70, 48);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(312, 21);
            this.cboTo.TabIndex = 129;
            this.cboTo.Text = "ua0413@zel.zantaz.com";
            // 
            // lnkBCC
            // 
            this.lnkBCC.Enabled = false;
            this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkBCC.Location = new System.Drawing.Point(30, 100);
            this.lnkBCC.Name = "lnkBCC";
            this.lnkBCC.Size = new System.Drawing.Size(36, 20);
            this.lnkBCC.TabIndex = 128;
            this.lnkBCC.TabStop = true;
            this.lnkBCC.Text = "BCC :";
            this.lnkBCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkCC
            // 
            this.lnkCC.Enabled = false;
            this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkCC.Location = new System.Drawing.Point(38, 76);
            this.lnkCC.Name = "lnkCC";
            this.lnkCC.Size = new System.Drawing.Size(28, 16);
            this.lnkCC.TabIndex = 127;
            this.lnkCC.TabStop = true;
            this.lnkCC.Text = "CC :";
            this.lnkCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkTo
            // 
            this.lnkTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkTo.Location = new System.Drawing.Point(18, 52);
            this.lnkTo.Name = "lnkTo";
            this.lnkTo.Size = new System.Drawing.Size(48, 16);
            this.lnkTo.TabIndex = 126;
            this.lnkTo.TabStop = true;
            this.lnkTo.Text = "To :";
            this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkFrom
            // 
            this.lnkFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFrom.Location = new System.Drawing.Point(18, 28);
            this.lnkFrom.Name = "lnkFrom";
            this.lnkFrom.Size = new System.Drawing.Size(48, 16);
            this.lnkFrom.TabIndex = 125;
            this.lnkFrom.TabStop = true;
            this.lnkFrom.Text = "From :";
            this.lnkFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // gbxDigiSafe
            // 
            this.gbxDigiSafe.Controls.Add(this.chkGUID);
            this.gbxDigiSafe.Controls.Add(this.cboFrom);
            this.gbxDigiSafe.Location = new System.Drawing.Point(6, 4);
            this.gbxDigiSafe.Name = "gbxDigiSafe";
            this.gbxDigiSafe.Size = new System.Drawing.Size(380, 140);
            this.gbxDigiSafe.TabIndex = 132;
            this.gbxDigiSafe.TabStop = false;
            // 
            // cboFrom
            // 
            this.cboFrom.Location = new System.Drawing.Point(64, 20);
            this.cboFrom.Name = "cboFrom";
            this.cboFrom.Size = new System.Drawing.Size(312, 21);
            this.cboFrom.TabIndex = 122;
            this.cboFrom.Text = "kcheung@zantaz.com";
            // 
            // chkModify
            // 
            this.chkModify.Checked = true;
            this.chkModify.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkModify.Location = new System.Drawing.Point(8, 208);
            this.chkModify.Name = "chkModify";
            this.chkModify.Size = new System.Drawing.Size(104, 16);
            this.chkModify.TabIndex = 147;
            this.chkModify.Text = "Modify Variable";
            this.chkModify.CheckedChanged += new System.EventHandler(this.chkModify_CheckedChanged);
            // 
            // ZelMsgPage
            // 
            this.Controls.Add(this.rdoSMTPClient);
            this.Controls.Add(this.chkModify);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.richBox);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.cboSMTP);
            this.Controls.Add(this.propGrid1);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.txtListFile);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.lnkListFile);
            this.Controls.Add(this.lnkFolder);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblSMTP);
            this.Controls.Add(this.cboCC);
            this.Controls.Add(this.cboBCC);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.lnkBCC);
            this.Controls.Add(this.lnkCC);
            this.Controls.Add(this.lnkTo);
            this.Controls.Add(this.lnkFrom);
            this.Controls.Add(this.gbxDigiSafe);
            this.Controls.Add(this.btnAbort);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.rdoMSAPI);
            this.Name = "ZelMsgPage";
            this.Size = new System.Drawing.Size(388, 428);
            this.Load += new System.EventHandler(this.ZelMsgPage_Load);
            this.gbxDigiSafe.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

        private void btnSend_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "ZelMsgPage.cs - btnSend_Click" );

            if( rdoMSAPI.Checked )
            {
                Debug.WriteLine("ZelMsgPage.cs - btnSend_Click : rdoMSAPI.Check");
                sendZelMailThread = new Thread( new ThreadStart(this.Thd_SendZelMail) );
                sendZelMailThread.Name = "sendZelMailThread";
                sendZelMailThread.Start();
                commObj.LogToFile( "Thread.log", "++" + sendZelMailThread.Name );
            }//end of if - send via MS API
            else
                if( rdoSMTPClient.Checked && !chkModify.Checked )
                {
                    Debug.WriteLine("ZelMsgPage.cs - btnSend_Click : rdoSMTPClient.Check ONLY");
                    sendSmtpMailThread = new Thread( new ThreadStart(this.Thd_SendSmtpMail) );
                    sendSmtpMailThread.Name = "sendSmtpMailThread";
                    sendSmtpMailThread.Start();
                    commObj.LogToFile( "Thread.log", "++" + sendSmtpMailThread.Name );
                }//end of if
            else
                if( rdoSMTPClient.Checked && chkModify.Checked )
                {
                    Debug.WriteLine("ZelMsgPage.cs - btnSend_Click : BOTH rdoSMTPClient & chkModify checked");
                    sendCustMailThread = new Thread( new ThreadStart(this.Thd_SendCustMail) );
                    sendCustMailThread.Name = "sendCustMailThread";
                    sendCustMailThread.Start();
                    commObj.LogToFile( "Thread.log", "++" + sendCustMailThread.Name );
                }//end of if

        }//end of btnSend_Click

        private void btnTest_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "QATool.cs - btnTest_Click" );
            this.Cursor = Cursors.WaitCursor;
            richBox.Text = commObj.TestSMTPConnection(cboSMTP.Text,cboPort.Text)?"Connection OK":"Connection FAIL";
            this.Cursor = Cursors.Default;		                
        }//end of btnTest_Click

        private void btnAbort_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "ZelMsgPage.cs - btnAbort_Click" );
            try
            {
                if( sendZelMailThread != null && sendZelMailThread.IsAlive )
                    KillSendZelMailThread();

                if( sendSmtpMailThread != null && sendSmtpMailThread.IsAlive )
                    KillSendSmtpMailThread();

                if( sendCustMailThread != null && sendCustMailThread.IsAlive )
                    KillSendCustMailThread();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine("ZelMsgPage.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                commObj.LogToFile("ZelMsgPage.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                MessageBox.Show( ex.Message + "\n" + ex.StackTrace, "Abort Exception" );
            }//end of catch        
        }//end of btnAbort_Click

        /// <summary>
        /// Initialize the property grid - attach the data object
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ZelMsgPage_Load(object sender, System.EventArgs e)
        {
            commObj.LogToFile( "ZelMsgPage.cs - ZELTool_Load: attach the data object to property grid" );
            propGrid1.SelectedObject = varObj;
        }// end of ZelMsgPage_Load

        /// <summary>
        /// Special case that cannot reuse SMTPSender....
        /// Open socket and stream a file into it... good for RFC822 file (eml file)
        /// Also modified the file value... this moment is $TO$ and $RUN$...
        /// </summary>
        public void Thd_SendCustMail()
        {
            Trace.WriteLine("QATool.cs - Thd_SendCustMail");
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;
            DateTime startTime = DateTime.Now;

            string strListLn = ""; // a line from list file
            StreamReader sr  = null;
            StreamReader sr2 = null;
            string strMsg;
            string strReply; // reply from smtp server           

            string _strFrom    = cboFrom.Text;
            string _strTo      = cboTo.Text;
            string _strServer  = cboSMTP.Text; // SMTP server name or ip
            string _strPortNum = cboPort.Text;

            TcpClient tcpClient = new TcpClient();
            try
            {
                string strTmp = ""; // line from input file (RFC822)
                sr = new StreamReader( txtListFile.Text );

                // Open smtp connection, then continue stream different file into socket
                Debug.WriteLine("\t Thd_SendCustMail() - Connect to smtp server");
                //                TcpClient tcpClient = new TcpClient();
                tcpClient.Connect( _strServer, int.Parse(_strPortNum) ); // inherit from System.Net.Sockets.TcpClient
                strReply = ReadFromSocket( tcpClient );
                if( strReply.Substring(0,3) != "220" )
                    throw new SMTPException( strReply ); // will catch in the caller (eg. main)

                Debug.WriteLine("\t Thd_SendCustMail() - Test the connection HELO");
                strMsg = "HELO world\r\n"; // test connection
                WriteToSocket( tcpClient, strMsg );
                strReply = ReadFromSocket( tcpClient );
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );

                while( (strListLn = sr.ReadLine()) != null ) 
                {
                    // Get the msg file name from list file 2nd col, pass it into HandleSendMail.
                    string [] splitStr = strListLn.Split( new Char [] {' '} );
                    Debug.WriteLine( "splitStr[1] == " + splitStr[1] ); // splitStr[1] is file name
                    if( splitStr[0] == "msg" ) // only handle line starting with msg
                    {
                        Debug.WriteLine("\t Thd_SendCustMail() - write mail from into socket");
                        strMsg = "MAIL FROM: " + _strFrom + "\r\n"; // set up 821 header
                        WriteToSocket( tcpClient, strMsg );
                        strReply = ReadFromSocket( tcpClient );
                        if( strReply.Substring(0,3) != "250" )
                            throw new SMTPException( strReply );

                        Debug.WriteLine("\t Thd_SendCustMail() - write rcpt to into socket");
                        strMsg = "RCPT TO: " + _strTo + "\r\n";
                        WriteToSocket( tcpClient, strMsg );
                        strReply = ReadFromSocket( tcpClient );
                        if( strReply.Substring(0, 3) != "250" )
                            throw new SMTPException( strReply );

                        Debug.WriteLine("\t Thd_SendCustMail() - Write DATA into socket - signaling SMTP server");
                        strMsg = "DATA\r\n";
                        WriteToSocket( tcpClient, strMsg );
                        strReply = ReadFromSocket( tcpClient );
                        if( strReply.Substring(0, 3) != "354" )
                            throw new SMTPException( strReply );

                        string filename = richBox.Text = txtFolder.Text + "\\" + splitStr[1];
                        int idxFirst;
                        int idxLast;

                        // read in mail
                        sr2 = new StreamReader( filename ); // file name
                        while( (strTmp = sr2.ReadLine()) != null )
                        {
                            // Take care "$$" - take out one.
                            strTmp = strTmp.Replace( "$$", "$" );
                            strTmp = Regex.Replace( strTmp, @"\$RUN\$", varObj.RUN, RegexOptions.IgnoreCase );
                            strTmp = Regex.Replace( strTmp, @"\$TO\$",  varObj.TO,  RegexOptions.IgnoreCase );

                            // Remove everything between the first '|' and last '|'
                            if( -1 != (idxFirst = strTmp.IndexOf('|')) ) // found frist
                            {
                                if( -1 != (idxLast = strTmp.LastIndexOf('|')) ) // found last
                                    if( idxFirst != idxLast ) // found two '|'
                                        strTmp = strTmp.Remove( idxFirst, idxLast-idxFirst+1 );
                            }//end of if - first index

                            // Remove everhthing between the first '{' and last '{'
                            if( -1 != (idxFirst = strTmp.IndexOf('{')) ) // found
                            {
                                if( -1 != (idxLast = strTmp.LastIndexOf('{')) )
                                    if( idxFirst != idxLast ) // found two
                                        strTmp = strTmp.Remove( idxFirst, idxLast-idxFirst+1 );
                            }//end of if - first index
                            WriteToSocket( tcpClient, strTmp );
                        }//end of while - loop through file
                        sr2.Close(); //close the stream reader

                        strMsg = "\r\n.\r\n"; // period - end of mail
                        Debug.WriteLine("\tSmtpSend() - end of mail: write dot into socket");
                        WriteToSocket( tcpClient, strMsg );

                        strReply = ReadFromSocket( tcpClient );
                        if( strReply.Substring(0, 3) != "250" )
                            throw new SMTPException( strReply );
                    }//end of if - only handle line starting with msg
                }//end of while - loop through list file 

                Debug.WriteLine("\tSmtpSend() - Send now and quit");
                strMsg = "QUIT\r\n"; // Send now...
                WriteToSocket( tcpClient, strMsg );

                Trace.WriteLine("\t Read tcpClient: " + tcpClient.ToString() );
                strReply = ReadFromSocket( tcpClient );
                if( strReply.IndexOf("221") == -1 )
                    throw new SMTPException( strReply );

//                tcpClient.Close(); // TCP connection - inherited from System.Net.Sockets.TcpClient

            }//end of try
            catch( SocketException ex )
            {
                Trace.WriteLine("\tSocket Exception: " + ex.Message.ToString() );

            }//end of catch - socket exception
            catch( IOException ioex )
            {
                Trace.WriteLine("\tIO Exception: " + ioex.Message.ToString() );
            }// end of catch - IO exception
            catch( QATool.SMTPException ex )
            {
                Trace.WriteLine("\tThd_SendCustMail() Exception: " + ex.SmtpMessage.ToString() );
                MessageBox.Show( ex.SmtpMessage.ToString(), "Thd_SendCustMail" );
            }//end of catch
            finally
            {
                Trace.WriteLine("\t Close TCP Connection ");
                tcpClient.Close(); // TCP connection - inherited from System.Net.Sockets.TcpClient

                if( sr != null )
                    sr.Close();

                if( sr2 != null )
                    sr2.Close();
            }//end of finally
            
            DateTime endTime = DateTime.Now;
            TimeSpan duration = endTime - startTime;
            string strTime = "Start time: " + startTime.ToString()
                + "\r\nEnd Time: " + endTime.ToString()
                + "\r\nDuration in Second: " + duration.TotalSeconds.ToString();
                            
            commObj.LogToFile( strTime );
            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;		

        }//end of Thd_SendCustMail

        /// <summary>
        /// Send ZEL mail when user click on the send button
        /// Generate in threading manner for better user experience
        /// </summary>
        public void Thd_SendSmtpMail()
        {
            Trace.WriteLine( "QATool.cs - Thd_SendSmtpMail" );

            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;
            DateTime startTime = DateTime.Now;

            string strListLn = ""; // a line from list file
            StreamReader sr = null;
            try
            {
                sr = new StreamReader( txtListFile.Text );   // Reading list file which contain msg file name
                while( (strListLn = sr.ReadLine()) != null ) // a line from List file
                {
                    QATool.SMTPSender smtpSender = new SMTPSender();
                    smtpSender.mailFrom    = cboFrom.Text;
                    smtpSender.mailTo      = cboTo.Text;
                    smtpSender.smtpServer  = cboSMTP.Text;
                    smtpSender.smtpPortNum = cboPort.Text;

                    // Get the msg file name from list file 2nd col, pass it into HandleSendMail.
                    string [] splitStr = strListLn.Split( new Char [] {' '} );

                    if( splitStr[0] == "msg" ) // only handle line starting with msg
                    {
                        Debug.WriteLine( "splitStr[1] == " + splitStr[1] );
                        richBox.Text = splitStr[1];
                        smtpSender.SmtpSend( txtFolder.Text + "\\" + splitStr[1] ); // after send mail, tcp connect close -> result delete object
                    }//end of if - only handle line starting with msg
                }//end of while
            }//end of try
            catch( QATool.SMTPException ex )
            {
                Trace.WriteLine("\tThd_SendSmtpMail() Exception: " + ex.SmtpMessage.ToString() );
                MessageBox.Show( ex.SmtpMessage.ToString(), "ZEL Message" );
            }//end of catch

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
        public void Thd_SendZelMail()
        {
            Trace.WriteLine( "QATool.cs - Thd_SendZelMail" );
			
            this.Cursor = Cursors.WaitCursor;
            btnSend.Enabled = false;
            DateTime startTime = DateTime.Now;

            string strListLn = ""; // a line from list file
            StreamReader sr = null;
            System.Net.Sockets.TcpClient tcpClient = null;
            try
            {		
                // If SmtpServer is not set, local SMTP server is used
                SmtpMail.SmtpServer = cboSMTP.Text;
                
                tcpClient = new TcpClient(); // open connection here - only open once
                tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboPort.Text) );			
                Debug.WriteLine("  +  Reading List file - " + txtListFile.Text);                
                try
                {
                    sr = new StreamReader( txtListFile.Text );   // Reading list file which contain msg file name
                    while( (strListLn = sr.ReadLine()) != null ) // a line from List file
                    {
                        // Get the msg file name from list file 2nd col, pass it into HandleSendMail.
                        string [] splitStr = strListLn.Split( new Char [] {' '} );

                        if( splitStr[0] == "msg" ) // only handle line starting with msg
                        {
                            Debug.WriteLine( "splitStr[1] == " + splitStr[1] );
                            HandleSendMail( splitStr[1] ); // 2nd col. index 1 - msg file name
                        }//end of if - only handle line starting with msg
                    }//end of while
                }//end of try
                catch( Exception ex )
                {
                    Trace.WriteLine( ex.Message.ToString() );
                    MessageBox.Show( ex.Message.ToString(), "Generic Exception" );
                }//end of catch
            }// end of try
            catch( System.Web.HttpException ex )
            {
                Trace.WriteLine( ex.Message.ToString() );
                MessageBox.Show( ex.Message.ToString(), "Generic HTTP Exception" );
            }//end of catch - generic exception
            finally
            {
                if( tcpClient != null )
                {
                    Debug.WriteLine( "Finally - close TCP Clinet connection");
                    commObj.LogToFile( "Finally - close TCP Clinet connection" );
                    tcpClient.Close();
                }//end of if 

                if( sr != null )
                {
                    Trace.WriteLine("Finally - close the Stream Reader");
                    commObj.LogToFile( "Finally - close the Stream Reader" );
                    sr.Close();
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
        }//end of Thd_SendZelMail

        /// <summary>
        /// Kill the send mail thread when program exit
        /// </summary>
        public void KillSendCustMailThread()
        {
            Trace.WriteLine("ZelMsgPage.cs - KillSendCustMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + sendCustMailThread.Name );
                sendCustMailThread.Abort(); // abort 
                sendCustMailThread.Join();  // ensure thread kill
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in Kill Thread:" + sendCustMailThread.Name );
            }//end of catch

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
        }//end of KillSendCustMailThread()

        /// <summary>
        /// Kill the send mail thread when program exit
        /// </summary>
        public void KillSendSmtpMailThread()
        {
            Trace.WriteLine("ZelMsgPage.cs - KillSendSmtpMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + sendSmtpMailThread.Name );
                sendSmtpMailThread.Abort(); // abort 
                sendSmtpMailThread.Join();  // ensure thread kill
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in Kill Thread:" + sendSmtpMailThread.Name );
            }//end of catch

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
        }//end of KillSendSmtpMailThread()

        /// <summary>
        /// Kill the send mail thread when program exit
        /// </summary>
        public void KillSendZelMailThread()
        {
            Trace.WriteLine("ZelMsgPage.cs - KillSendZelMailThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "++ Kill Thread:" + sendZelMailThread.Name );
                sendZelMailThread.Abort(); // abort 
                sendZelMailThread.Join();  // ensure thread kill
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Thread.log", "\t Exception ocurr in KillZelSendMailThread:" + sendZelMailThread.Name );
            }//end of catch

            this.Cursor = Cursors.Default;
            btnSend.Enabled = true;
        }//end of KillSendMailThread

        public void HandleSendMail( String strInFile )
        {
            Trace.WriteLine( "ZELTool.cs - HandleSendMail" );

            int    counter = 0; // mail sent counter
            string strGUID = "";
            string strMsgFile = txtFolder.Text + "\\" + strInFile;

            // TO DO: Check does file exist? No, return

            StreamReader sr = null;
            try
            {
                counter++;
                Debug.WriteLine( "\t - HandleSendMail - inside while loop : " + counter.ToString() );

                richBox.Text = "\r\n- Read line : ";
                if( chkGUID.Checked )
                {
                    strGUID = System.Guid.NewGuid().ToString();
                    txtSubject.Text = counter.ToString() + " " + strGUID;
                }//end of if - GUID			

                MailMessage mailMsg = new MailMessage();

                mailMsg.From	= cboFrom.Text;	// single line from GUI
                mailMsg.To		= cboTo.Text;	// single line from GUI
                mailMsg.Cc		= cboCC.Text;	// user input, may != To
                mailMsg.Bcc		= cboBCC.Text;	// user input, may != To
                mailMsg.Subject = txtSubject.Text;
                mailMsg.Body    = BuildMsgBody( strMsgFile ) + "/r/nKentest/r/n";


                Debug.WriteLine( "mailMsg.Body == " + mailMsg.Body );
#if(DEBUG)
                commObj.LogToFile("msgBody.txt", mailMsg.Body);
#endif

                try
                {
                    Debug.WriteLine("  +  HandleSendMail - send mail la");
                    richBox.Text = "Do the send mail.\r\n+ ZEL Msg Sent Info " + strInFile;
                    SmtpMail.Send( mailMsg );

//                    Thread.Sleep( 500 ); // kentest
#if(DEBUG)
                    commObj.LogToFile("msgBody.txt", "Finished Send mail la\n");
#endif
                }//end of try
                catch( System.Web.HttpException ex )
                {
                    Trace.WriteLine(ex.Message.ToString());
                    commObj.LogToFile( "Exception.LOG", ex.Message.ToString() + "\n" + ex.StackTrace );
                }// end of catch

                //save                commObj.LogToFile( richBox.Text ); // save into log
            }//end of try - IOException
            catch( Exception ex )
            {
                Trace.WriteLine("Exception" + ex.Message.ToString());
                commObj.LogToFile( "Exception.LOG", ex.Message.ToString() + "\n" + ex.StackTrace );
            }//end of catch - IOException	
            finally
            {
                if( sr != null )
                {
                    Trace.WriteLine("Finally - close the Stream Reader");
                    sr.Close();
                }//end of if
            }//end of finally
        }// end of HandleSendMail

        private void lnkListFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "ZELTool.cs - lnkListFile_LinkClicked" );

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.ShowReadOnly = true;
            ofDlg.RestoreDirectory = false;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtListFile.Text = m_listFile   = ofDlg.FileName;
                // get the data folder path without ending '\'
                txtFolder.Text   = m_dataFolder = m_listFile.Substring( 0, m_listFile.LastIndexOf('\\') );
            }//end of if
        }//end of lnkListFile_LinkClicked

        private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "QATool.cs - lnkFolder_LinkClicked" );
            FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
            if( txtFolder.Text != null )
                fbDlg.SelectedPath = txtFolder.Text;  // set the default folder

            if( fbDlg.ShowDialog() == DialogResult.OK )
            {
                txtFolder.Text = m_dataFolder = fbDlg.SelectedPath;
            }
        }//end of HandleSendMail

        /// <summary>
        /// 1) Trim $$
        /// 2) Replace $RUN$ or $run$ with user input variable
        /// 3) Replace $TO$ with dummy variable
        /// 4) Trim EVERYTHING between first '|' and last '|'
        /// 5) Trim EVERYTHING between frist '{' and last '{'
        /// </summary>
        /// <param name="lstFileName">Full path file name of list file</param>
        /// <returns>Body Message CHECK: what is the max size for string?</returns>
        public StringBuilder BuildMsgBody( string msgFile )
        {
            string strTmp = "";
            StringBuilder strMsg = new StringBuilder();
            StreamReader sr = null;
            try
            {
                sr = new StreamReader( msgFile ); 
                while( (strTmp = sr.ReadLine()) != null ) 
                {
                    int idxFirst;
                    int idxLast;
									
                    if( chkModify.Checked ) // do the modification
                    {
                        // Take care "$$" - take out one.
                        strTmp = strTmp.Replace( "$$", "$" );
                        strTmp = Regex.Replace( strTmp, @"\$RUN\$", varObj.RUN, RegexOptions.IgnoreCase );
                        strTmp = Regex.Replace( strTmp, @"\$TO\$",  varObj.TO,  RegexOptions.IgnoreCase );

                        // Remove everything between the first '|' and last '|'
                        if( -1 != (idxFirst = strTmp.IndexOf('|')) ) // found frist
                        {
                            if( -1 != (idxLast = strTmp.LastIndexOf('|')) ) // found last
                                if( idxFirst != idxLast ) // found two '|'
                                    strTmp = strTmp.Remove( idxFirst, idxLast-idxFirst+1 );
                        }//end of if - first index

                        // Remove everhthing between the first '{' and last '{'
                        if( -1 != (idxFirst = strTmp.IndexOf('{')) ) // found
                        {
                            if( -1 != (idxLast = strTmp.LastIndexOf('{')) )
                                if( idxFirst != idxLast ) // found two
                                    strTmp = strTmp.Remove( idxFirst, idxLast-idxFirst+1 );
                        }//end of if - first index
                    }//end of if - modify checked

                    strMsg = strMsg.Append(strTmp);
                    Debug.WriteLine( "strMsg == " + strMsg );

                }//end of while
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine(ex.Message.ToString());
                commObj.LogToFile( "Exception.LOG", ex.Message.ToString() + "\n" + ex.StackTrace );
            }//end of catch - IOException	
            finally
            {
                if( sr != null )
                {
                    Trace.WriteLine("Finally - close the Stream Reader");
                    sr.Close();
                }//end of if
            }//end of finally

            return strMsg;
        }//end of BuildMsgBody

        private void chkModify_CheckedChanged(object sender, System.EventArgs e)
        {
            if( chkModify.Checked )
            {
                propGrid1.Enabled = true;
                propGrid1.ViewForeColor = Color.Black;
            }
            else
            {
                propGrid1.Enabled = false;
                propGrid1.ViewForeColor = Color.DimGray;
            }
        }//end of chkModify_CheckedChanged

        /// <summary>
        /// Write data to socket in ASCII format. dotNet string class is unicode. ie: need to convert to ASCII encoding.
        /// </summary>
        public void WriteToSocket( TcpClient tcp, string msg )
        {
            Trace.WriteLine("ZelMsgPage.cs - WriteToSocket():" + msg.ToString() );
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            byte[] writeBuffer = new byte[BYTE_SIZE]; // 8K
            writeBuffer = asciiEncoding.GetBytes( msg );

            NetworkStream nwStream = tcp.GetStream();
            nwStream.Write( writeBuffer, 0, writeBuffer.Length );

        }//end of WriteToSocket

        /// <summary>
        /// Read data stream from socket and convert ASCII back to native dotNet string.
        /// </summary>
        /// <returns></returns>
        public string ReadFromSocket( TcpClient tcp )
        {
//            Trace.WriteLine( "ZelMsgPage.cs - ReadFromSocket()" );
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            byte[] serverBuffer = new byte[BYTE_SIZE]; // 8K

            NetworkStream nwStream = tcp.GetStream();
            int count = nwStream.Read( serverBuffer, 0, serverBuffer.Length );
            if( count == 0 ) // no more data
                return( "" );
            else
                return( asciiEncoding.GetString(serverBuffer, 0, count) );
        }//end of ReadFromSocket

        /// <summary>
        /// Mainly handle enable or disable the UI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rdoMSAPI_Click(object sender, System.EventArgs e)
        {
            lnkCC.Enabled       = true;
            lnkBCC.Enabled      = true;
            cboCC.Enabled       = true;
            cboBCC.Enabled      = true;
            txtSubject.Enabled  = true;
            chkGUID.Enabled     = true;
        
        }//end of rdoMSAPI_Click

        private void rdoSMTPClient_Click(object sender, System.EventArgs e)
        {
            lnkCC.Enabled       = false;
            lnkBCC.Enabled      = false;
            cboCC.Enabled       = false;
            cboBCC.Enabled      = false;
            txtSubject.Enabled  = false;
            chkGUID.Enabled     = false;
        
        }

	}
}
 