using System;
using System.Collections;
using System.ComponentModel;
using System.Data.SqlClient;
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
	/// Summary description for DBMailPage.
	/// </summary>
	public class DBMailPage : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Label lblIdPwd;
		private System.Windows.Forms.TextBox txtUID;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.ToolTip ttpDBMail;
		private System.Windows.Forms.Label lblSQLServer;
		private System.Windows.Forms.DataGrid dataGrid;
		private System.Windows.Forms.ComboBox cboSqlHostName;
		private System.Windows.Forms.Button btnViewData;
		private System.Windows.Forms.TextBox txtFolder;
		private System.Windows.Forms.LinkLabel lnkFolder;
		private System.Windows.Forms.Button btnSend;
		private System.Windows.Forms.CheckBox chkAttach;
		private System.Windows.Forms.ComboBox cboSMTPPort;
		private System.Windows.Forms.ComboBox cboSMTP;
		private System.Windows.Forms.Label lblPort;
		private System.Windows.Forms.Label lblSMTP;
		private System.Windows.Forms.ComboBox cboSQLPort;
		private System.Windows.Forms.Label lblPort2;
		private System.Windows.Forms.TextBox txtOutput;
		private System.Windows.Forms.Button btnConnect;
		private System.Windows.Forms.Button btnTest;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ComboBox cboDBName;
		private System.Windows.Forms.Label lblDBName;
		private System.Windows.Forms.Label lblTo;
		private System.Windows.Forms.ComboBox cboTo;
		private System.Windows.Forms.Label lblSubject;
		private System.Windows.Forms.TextBox txtSubject;

		// custom declaration
		private QATool.CommObj commObj = new CommObj();
		private System.Windows.Forms.Label lblTableName;
		private System.Windows.Forms.ComboBox cboTableName;
		private Thread sendDBMailThread; // make it global -> more control for start and abort

		public DBMailPage()
		{
            Debug.WriteLine("DBMailPage.cs - Initialize DBMailPage Object");
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
			commObj.InitComboBoxItem( cboTo, "[To Address]" );
			commObj.InitComboBoxItem( cboSMTP, "[SMTP IP]" );
			commObj.InitComboBoxItem( cboSMTPPort, "[Port]" );
			commObj.InitComboBoxItem( cboSqlHostName, "[SQL IP]" );
			commObj.InitComboBoxItem( cboSQLPort, "[Port]" );
		}// end of constructor

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
                Debug.WriteLine("DBMailPage.cs - Deposing DBMailPage Object");

				if( (sendDBMailThread != null) && (sendDBMailThread.IsAlive) )
					this.KillSendDBMailThread();

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
            this.lblIdPwd = new System.Windows.Forms.Label();
            this.txtUID = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.ttpDBMail = new System.Windows.Forms.ToolTip(this.components);
            this.btnConnect = new System.Windows.Forms.Button();
            this.cboSqlHostName = new System.Windows.Forms.ComboBox();
            this.cboDBName = new System.Windows.Forms.ComboBox();
            this.btnViewData = new System.Windows.Forms.Button();
            this.cboSMTPPort = new System.Windows.Forms.ComboBox();
            this.cboSMTP = new System.Windows.Forms.ComboBox();
            this.cboSQLPort = new System.Windows.Forms.ComboBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.btnTest = new System.Windows.Forms.Button();
            this.cboTo = new System.Windows.Forms.ComboBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.cboTableName = new System.Windows.Forms.ComboBox();
            this.lblSQLServer = new System.Windows.Forms.Label();
            this.dataGrid = new System.Windows.Forms.DataGrid();
            this.lblDBName = new System.Windows.Forms.Label();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.chkAttach = new System.Windows.Forms.CheckBox();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblSMTP = new System.Windows.Forms.Label();
            this.lblPort2 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.lblTo = new System.Windows.Forms.Label();
            this.lblSubject = new System.Windows.Forms.Label();
            this.lblTableName = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // lblIdPwd
            // 
            this.lblIdPwd.Location = new System.Drawing.Point(4, 12);
            this.lblIdPwd.Name = "lblIdPwd";
            this.lblIdPwd.Size = new System.Drawing.Size(76, 12);
            this.lblIdPwd.TabIndex = 0;
            this.lblIdPwd.Text = "Id / Password";
            this.lblIdPwd.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtUID
            // 
            this.txtUID.Location = new System.Drawing.Point(80, 8);
            this.txtUID.Name = "txtUID";
            this.txtUID.Size = new System.Drawing.Size(124, 20);
            this.txtUID.TabIndex = 1;
            this.txtUID.Text = "sa";
            this.ttpDBMail.SetToolTip(this.txtUID, "DB Login Name");
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(208, 8);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '+';
            this.txtPassword.Size = new System.Drawing.Size(172, 20);
            this.txtPassword.TabIndex = 2;
            this.txtPassword.Text = "marketgoat";
            this.ttpDBMail.SetToolTip(this.txtPassword, "Re-type overwrite default");
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(318, 32);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(64, 20);
            this.btnConnect.TabIndex = 5;
            this.btnConnect.Text = "Connect";
            this.ttpDBMail.SetToolTip(this.btnConnect, "Test Connection");
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // cboSqlHostName
            // 
            this.cboSqlHostName.Location = new System.Drawing.Point(80, 32);
            this.cboSqlHostName.Name = "cboSqlHostName";
            this.cboSqlHostName.Size = new System.Drawing.Size(124, 21);
            this.cboSqlHostName.TabIndex = 7;
            this.cboSqlHostName.Text = "10.1.31.101";
            this.ttpDBMail.SetToolTip(this.cboSqlHostName, "SQL Server IP or server name");
            // 
            // cboDBName
            // 
            this.cboDBName.Items.AddRange(new object[] {
                                                           "pubs"});
            this.cboDBName.Location = new System.Drawing.Point(80, 56);
            this.cboDBName.Name = "cboDBName";
            this.cboDBName.Size = new System.Drawing.Size(124, 21);
            this.cboDBName.TabIndex = 9;
            this.cboDBName.Text = "SearchData";
            this.ttpDBMail.SetToolTip(this.cboDBName, "Table name");
            // 
            // btnViewData
            // 
            this.btnViewData.Location = new System.Drawing.Point(318, 56);
            this.btnViewData.Name = "btnViewData";
            this.btnViewData.Size = new System.Drawing.Size(64, 21);
            this.btnViewData.TabIndex = 10;
            this.btnViewData.Text = "View Data";
            this.ttpDBMail.SetToolTip(this.btnViewData, "Load data into Data Gride");
            this.btnViewData.Click += new System.EventHandler(this.btnViewData_Click);
            // 
            // cboSMTPPort
            // 
            this.cboSMTPPort.ItemHeight = 13;
            this.cboSMTPPort.Location = new System.Drawing.Point(244, 80);
            this.cboSMTPPort.Name = "cboSMTPPort";
            this.cboSMTPPort.Size = new System.Drawing.Size(72, 21);
            this.cboSMTPPort.Sorted = true;
            this.cboSMTPPort.TabIndex = 47;
            this.cboSMTPPort.Text = "25";
            this.ttpDBMail.SetToolTip(this.cboSMTPPort, "Hard Coded");
            // 
            // cboSMTP
            // 
            this.cboSMTP.ItemHeight = 13;
            this.cboSMTP.Items.AddRange(new object[] {
                                                         ""});
            this.cboSMTP.Location = new System.Drawing.Point(80, 80);
            this.cboSMTP.Name = "cboSMTP";
            this.cboSMTP.Size = new System.Drawing.Size(124, 21);
            this.cboSMTP.Sorted = true;
            this.cboSMTP.TabIndex = 46;
            this.cboSMTP.Text = "10.1.89.201";
            this.ttpDBMail.SetToolTip(this.cboSMTP, "Read From QATool.ini");
            // 
            // cboSQLPort
            // 
            this.cboSQLPort.ItemHeight = 13;
            this.cboSQLPort.Location = new System.Drawing.Point(244, 32);
            this.cboSQLPort.Name = "cboSQLPort";
            this.cboSQLPort.Size = new System.Drawing.Size(72, 21);
            this.cboSQLPort.Sorted = true;
            this.cboSQLPort.TabIndex = 55;
            this.cboSQLPort.Text = "1433";
            this.ttpDBMail.SetToolTip(this.cboSQLPort, "SQL server port Number");
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(318, 104);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(64, 21);
            this.btnSend.TabIndex = 52;
            this.btnSend.Text = "Send";
            this.ttpDBMail.SetToolTip(this.btnSend, "Send batch e-mail from DB");
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(318, 80);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(64, 21);
            this.btnTest.TabIndex = 48;
            this.btnTest.Text = "Test";
            this.ttpDBMail.SetToolTip(this.btnTest, "Test SMTP Connection");
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // cboTo
            // 
            this.cboTo.ItemHeight = 13;
            this.cboTo.Location = new System.Drawing.Point(80, 104);
            this.cboTo.Name = "cboTo";
            this.cboTo.Size = new System.Drawing.Size(236, 21);
            this.cboTo.Sorted = true;
            this.cboTo.TabIndex = 59;
            this.cboTo.Text = "login0@company1.zantaz.com";
            this.ttpDBMail.SetToolTip(this.cboTo, "Read from QATool.ini");
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(80, 128);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(236, 20);
            this.txtSubject.TabIndex = 61;
            this.txtSubject.Text = "DB";
            this.ttpDBMail.SetToolTip(this.txtSubject, "For limit the return search result");
            // 
            // cboTableName
            // 
            this.cboTableName.ItemHeight = 13;
            this.cboTableName.Items.AddRange(new object[] {
                                                              "dom1addr",
                                                              "dom2addr",
                                                              "dom3addr",
                                                              "MailData"});
            this.cboTableName.Location = new System.Drawing.Point(244, 56);
            this.cboTableName.Name = "cboTableName";
            this.cboTableName.Size = new System.Drawing.Size(72, 21);
            this.cboTableName.Sorted = true;
            this.cboTableName.TabIndex = 63;
            this.cboTableName.Text = "MailData";
            this.ttpDBMail.SetToolTip(this.cboTableName, "Table Name");
            // 
            // lblSQLServer
            // 
            this.lblSQLServer.Location = new System.Drawing.Point(4, 36);
            this.lblSQLServer.Name = "lblSQLServer";
            this.lblSQLServer.Size = new System.Drawing.Size(76, 12);
            this.lblSQLServer.TabIndex = 3;
            this.lblSQLServer.Text = "SQL Server";
            this.lblSQLServer.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dataGrid
            // 
            this.dataGrid.DataMember = "";
            this.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid.Location = new System.Drawing.Point(4, 176);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.ReadOnly = true;
            this.dataGrid.Size = new System.Drawing.Size(380, 220);
            this.dataGrid.TabIndex = 6;
            // 
            // lblDBName
            // 
            this.lblDBName.Location = new System.Drawing.Point(4, 60);
            this.lblDBName.Name = "lblDBName";
            this.lblDBName.Size = new System.Drawing.Size(76, 12);
            this.lblDBName.TabIndex = 8;
            this.lblDBName.Text = "DB Name";
            this.lblDBName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtFolder
            // 
            this.txtFolder.Enabled = false;
            this.txtFolder.Location = new System.Drawing.Point(148, 152);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(168, 20);
            this.txtFolder.TabIndex = 54;
            this.txtFolder.Text = "c:\\TestData";
            this.ttpDBMail.SetToolTip(this.txtFolder, "Full path of folder that contains attachment file");
            // 
            // lnkFolder
            // 
            this.lnkFolder.Enabled = false;
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.Location = new System.Drawing.Point(96, 156);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(42, 16);
            this.lnkFolder.TabIndex = 53;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpDBMail.SetToolTip(this.lnkFolder, "Select folder that contains attachment files");
            this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
            // 
            // chkAttach
            // 
            this.chkAttach.Location = new System.Drawing.Point(6, 156);
            this.chkAttach.Name = "chkAttach";
            this.chkAttach.Size = new System.Drawing.Size(82, 16);
            this.chkAttach.TabIndex = 51;
            this.chkAttach.Text = "Attachment";
            this.ttpDBMail.SetToolTip(this.chkAttach, "Include attachment");
            this.chkAttach.CheckedChanged += new System.EventHandler(this.chkAttach_CheckedChanged);
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(208, 84);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(36, 16);
            this.lblPort.TabIndex = 45;
            this.lblPort.Text = "Port #";
            // 
            // lblSMTP
            // 
            this.lblSMTP.Location = new System.Drawing.Point(4, 84);
            this.lblSMTP.Name = "lblSMTP";
            this.lblSMTP.Size = new System.Drawing.Size(76, 16);
            this.lblSMTP.TabIndex = 44;
            this.lblSMTP.Text = "SMTP Server";
            this.lblSMTP.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPort2
            // 
            this.lblPort2.Location = new System.Drawing.Point(208, 36);
            this.lblPort2.Name = "lblPort2";
            this.lblPort2.Size = new System.Drawing.Size(36, 16);
            this.lblPort2.TabIndex = 56;
            this.lblPort2.Text = "Port #";
            this.lblPort2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(4, 400);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ReadOnly = true;
            this.txtOutput.Size = new System.Drawing.Size(380, 20);
            this.txtOutput.TabIndex = 57;
            this.txtOutput.Text = "txtOutput";
            // 
            // lblTo
            // 
            this.lblTo.Location = new System.Drawing.Point(4, 108);
            this.lblTo.Name = "lblTo";
            this.lblTo.Size = new System.Drawing.Size(76, 16);
            this.lblTo.TabIndex = 58;
            this.lblTo.Text = "To :";
            this.lblTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblSubject
            // 
            this.lblSubject.Location = new System.Drawing.Point(4, 132);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(76, 16);
            this.lblSubject.TabIndex = 60;
            this.lblSubject.Text = "Subject:";
            this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTableName
            // 
            this.lblTableName.Location = new System.Drawing.Point(208, 60);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(36, 16);
            this.lblTableName.TabIndex = 62;
            this.lblTableName.Text = "Table";
            this.lblTableName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // DBMailPage
            // 
            this.Controls.Add(this.cboTableName);
            this.Controls.Add(this.lblTableName);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.cboTo);
            this.Controls.Add(this.lblTo);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.lblPort2);
            this.Controls.Add(this.cboSQLPort);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.lnkFolder);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.chkAttach);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.cboSMTPPort);
            this.Controls.Add(this.cboSMTP);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblSMTP);
            this.Controls.Add(this.btnViewData);
            this.Controls.Add(this.cboDBName);
            this.Controls.Add(this.lblDBName);
            this.Controls.Add(this.cboSqlHostName);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.lblSQLServer);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUID);
            this.Controls.Add(this.lblIdPwd);
            this.Name = "DBMailPage";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

		private void btnConnect_Click(object sender, System.EventArgs e)
		{            
			string strConnect =	"Persist Security Info=False;"
				              + "database=" + cboDBName.Text
				              + "; server=" + cboSqlHostName.Text + ' ' + cboSQLPort.Text
				              + "; uid=" + txtUID.Text
				              + "; pwd=" + txtPassword.Text;

			this.Cursor = Cursors.WaitCursor;
			txtOutput.Text = commObj.TestSQLConnection(strConnect)?"SQL Connection OK":"SQL Connection FAIL";
			this.Cursor = Cursors.Default;		
		}//end of btnConnect_Click

		/// <summary>
		/// Test the SMTP Connection by calling a method in common object
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnTest_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			txtOutput.Text = commObj.TestSMTPConnection(cboSMTP.Text, cboSMTPPort.Text)?"SMTP Connection OK":"SMTP Connection FAIL";
			this.Cursor = Cursors.Default;		
		}//end of btnTest_Click

		private void btnViewData_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "DBMailPage - btnConnect_Click" );
			Debug.WriteLine( "User = " + txtUID.Text + " Password = " + txtPassword.Text );
			// connect to local server - northwind db
			string connectStr =	"Persist Security Info=False;"
   							  + "database=" + cboDBName.Text
							  + "; server=" + cboSqlHostName.Text + ' ' + cboSQLPort.Text
							  + "; uid=" + txtUID.Text
							  + "; pwd=" + txtPassword.Text;
			
			// get records from the MailData table - hard coded
//			string commandStr = "Select GUID, FromField from MailData";
			string commandStr = "Select GUID, FromField from " + cboTableName.Text;

			try
			{
				// create the data set command object and the DataSet
				DataSet ds = new DataSet();
				SqlDataAdapter DataAdapter = new SqlDataAdapter( commandStr, connectStr );

				// fill the data set
				DataAdapter.Fill( ds, cboTableName.Text ); //Fill the data set object
				dataGrid.DataSource = ds.Tables[cboTableName.Text].DefaultView;

				DataAdapter.Dispose(); // delete component resource after loading the data
			}
			catch( Exception ex )
			{
				Trace.WriteLine( "DB Exception occur ", ex.Message.ToString() );
				MessageBox.Show( ex.Message.ToString(), "Error" );
			}//end of catch 			
		}// end of btnViewData_Click

		private void btnSend_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "DBMailPage.cs - btnSend_Click" );

			sendDBMailThread = new Thread( new ThreadStart(this.Thd_SendDBMail) );
			sendDBMailThread.Name = "dbMailThread";
			sendDBMailThread.Start();

            commObj.LogToFile( "Thread.log", "++ dbMailThread Start ++" );
		}// end of btnSend_Click

		private void chkAttach_CheckedChanged(object sender, System.EventArgs e)
		{
			if( chkAttach.Checked )
			{
				lnkFolder.Enabled = true;
				txtFolder.Enabled = true;
			}//end of if
			else
			{
				lnkFolder.Enabled = false;
				txtFolder.Enabled = false;
			}//end of else
		}//end of chkAttach_CheckedChanged

		private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
            if( txtFolder.Text != null )
                fbDlg.SelectedPath = txtFolder.Text;  // set the default folder

			if( fbDlg.ShowDialog() == DialogResult.OK )
			{
				txtFolder.Text = fbDlg.SelectedPath;
			}		
		}//end of lnkFolder_LinkClicked

		/// <summary>
		/// Read from the DB and then construct the mail -> send it.
		/// Also generate in threading maner for better user feel
		/// </summary>
		public void Thd_SendDBMail()
		{
			Trace.WriteLine( "DBMailPage.cs - Thd_SendDBMail" );
			Debug.WriteLine( "User = " + txtUID.Text + " Password = " + txtPassword.Text );
			
			this.Cursor = Cursors.WaitCursor;
			btnSend.Enabled = false;

			int count = 0;		// count number of mail from DB
			// connect to local server - northwind db
			string connectStr =	"Persist Security Info=False;"
				+ "database=" + cboDBName.Text
				+ "; server=" + cboSqlHostName.Text + ' ' + cboSQLPort.Text
				+ "; uid=" + txtUID.Text
				+ "; pwd=" + txtPassword.Text;
			
			// get records from the MailData table - hard coded
//			string commandStr = "Select GUID, FromField, Attachment from MailData";
			string commandStr = "Select GUID, FromField, Attachment from " + cboTableName.Text;

			try
			{
				// create the data set command object and the DataSet
				DataSet ds = new DataSet();
				SqlDataAdapter DataAdapter = new SqlDataAdapter( commandStr, connectStr );

//				DataAdapter.Fill( ds, "MailData" ); //Fill the data set object
				DataAdapter.Fill( ds, cboTableName.Text );
				DataTable dataTable = ds.Tables[0]; // Get the one table from the DataSet
				
				try
				{			
					MailMessage mailMsg = new MailMessage();
					foreach( DataRow dataRow in dataTable.Rows )
					{
						mailMsg.From    = dataRow["FromField"].ToString();
						mailMsg.To      = cboTo.Text;
						//						mailMsg.Cc      = txtCC.Text;
						//						mailMsg.Bcc     = txtBCC.Text;
						mailMsg.Subject = txtSubject.Text + " " + dataRow["GUID"].ToString();
						mailMsg.Body    = "Body: " + DateTime.Now + "\n"
							+ dataRow["FromField"].ToString() + " "
							+ dataRow["GUID"].ToString()
							+ "Mail count is " + count++;

						if( chkAttach.Checked )
						{						
							txtFolder.Text.TrimStart( new char[] {' '} ); // trim leading space		
							String strFileFullName = txtFolder.Text + "\\" + dataRow["Attachment"];;
							mailMsg.Attachments.Clear(); // clear the attachment list
							mailMsg.Attachments.Add( new MailAttachment( strFileFullName, MailEncoding.Base64 ) );							

							mailMsg.Body += "\nAttachment Include";
						}//end of if

						// If SmtpServer is not set, local SMTP server is used
						SmtpMail.SmtpServer = cboSMTP.Text;

						// Test the connection - if smtp server down, exit.
						System.Net.Sockets.TcpClient tcpClient = new TcpClient();
						tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboSMTPPort.Text) );

						txtOutput.Text = "Sending " + mailMsg.From;
						SmtpMail.Send( mailMsg );

						Trace.WriteLine( "\t - Message Sent: " + cboTo.Text );
						tcpClient.Close();
					}//end of foreach - DataRow								
				}// end of try
				catch( System.Web.HttpException ex )
				{
					Trace.WriteLine( "MailPage.cs - HTTP Exception" );
					MessageBox.Show(ex.Message.ToString());
				}//end of catch				
			}//end of try
			catch( Exception ex )
			{
				Trace.WriteLine( "DB Exception occur ", ex.Message.ToString() );
				MessageBox.Show( ex.Message.ToString(), "Error" );
			}// end of catch
			this.Cursor = Cursors.Default;
			btnSend.Enabled = true;
			txtOutput.Text = "End of Sending mail...";		
		}// end of Thd_SendDBMail

		/// <summary>
		/// Kill the send db mail when program exit
		/// </summary>
		public void KillSendDBMailThread()
		{
			Trace.WriteLine("DBMailPage.cs - KillSendDBMailThread()");
			try
			{
                commObj.LogToFile( "Thread.log", "++Kill Thread:"+sendDBMailThread.Name);
				sendDBMailThread.Abort(); // abort
                sendDBMailThread.Join();  // require for ensure the thread kill
			}//end of try 
			catch( ThreadAbortException thdEx )
			{
				Trace.WriteLine( thdEx.Message );
				txtOutput.Text += "\nAborting the send DB mail thread";
			}//end of catch		
		}

	}
}
