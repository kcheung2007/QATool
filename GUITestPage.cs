using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.IO;
using System.Text;
using System.Threading;
using System.Web;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for GUITestPage.
	/// </summary>
	public class GUITestPage : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Label lblUid;
		private System.Windows.Forms.Label lblPassword;
		private System.Windows.Forms.ComboBox cboUid;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.LinkLabel lnkGuidFile;
		private System.Windows.Forms.TextBox txtGuidFile;
		private System.Windows.Forms.Button btnLogin;
		private System.Windows.Forms.RichTextBox rtbDisplay;
		private System.Windows.Forms.ToolTip ttpGUITest;
		private System.Windows.Forms.Label lblURL;
		private System.Windows.Forms.ComboBox cboURL;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton rdoGUIDFile;
		private System.Windows.Forms.Button btnSearch;
		private System.Windows.Forms.Label lblToken;
		private System.Windows.Forms.RadioButton rdoToken;
		private System.Windows.Forms.TextBox txtToken;

		private URLDataObj urlData = new URLDataObj();
        private System.Windows.Forms.DateTimePicker dtpStartDate;
        private System.Windows.Forms.DateTimePicker dtpEndDate;
        private System.Windows.Forms.Label lblStartDate;
        private System.Windows.Forms.Label lblEndDate;
        private QATool.CommObj commObj = new CommObj();
		private System.Windows.Forms.TextBox txtQueryFile;
		private System.Windows.Forms.RadioButton rdoQueryFile;
		private System.Windows.Forms.TextBox txtSubject;
		private System.Windows.Forms.CheckBox chkOption;
		private System.Windows.Forms.LinkLabel lnkQueryFile;

		private const int CHECK_SIZE = 50;

		private System.Windows.Forms.NumericUpDown nupWait;
		private System.Windows.Forms.Label lblSec;

		private Thread searchThread; // make it global -> more control for start and abort

		static private int m_found = 0;
		private System.Windows.Forms.Button btnStop;
		private System.Windows.Forms.Label lblRepository;
		private System.Windows.Forms.ComboBox cboRepository; // number of message found
		static private bool m_login = false; // default

		public GUITestPage()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            Debug.WriteLine("GUITestPage.cs - Initial GUI Test page");

			// TODO: Add any initialization after the InitializeComponent call
			commObj.InitComboBoxItem( cboUid, "[Login ID]" );
			commObj.InitComboBoxItem( cboURL, "[Login URL]" );
			commObj.InitComboBoxItem( cboRepository, "[Repository]" );
		}//end of constructor

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
                Debug.WriteLine("GUITestPage.cs - Dispose GUI Page");

				if( m_login )
					HandleLogoff(); // no matter what, run it...

				if( (searchThread != null) && searchThread.IsAlive )
					KillSearchThread();

				if(components != null)
				{
					Debug.WriteLine("\t Dispose component");
					components.Dispose();                    
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
			this.lblUid = new System.Windows.Forms.Label();
			this.lblPassword = new System.Windows.Forms.Label();
			this.cboUid = new System.Windows.Forms.ComboBox();
			this.txtPassword = new System.Windows.Forms.TextBox();
			this.lnkGuidFile = new System.Windows.Forms.LinkLabel();
			this.txtGuidFile = new System.Windows.Forms.TextBox();
			this.btnLogin = new System.Windows.Forms.Button();
			this.rtbDisplay = new System.Windows.Forms.RichTextBox();
			this.ttpGUITest = new System.Windows.Forms.ToolTip(this.components);
			this.cboURL = new System.Windows.Forms.ComboBox();
			this.txtToken = new System.Windows.Forms.TextBox();
			this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
			this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
			this.rdoToken = new System.Windows.Forms.RadioButton();
			this.rdoGUIDFile = new System.Windows.Forms.RadioButton();
			this.txtQueryFile = new System.Windows.Forms.TextBox();
			this.rdoQueryFile = new System.Windows.Forms.RadioButton();
			this.txtSubject = new System.Windows.Forms.TextBox();
			this.chkOption = new System.Windows.Forms.CheckBox();
			this.lnkQueryFile = new System.Windows.Forms.LinkLabel();
			this.nupWait = new System.Windows.Forms.NumericUpDown();
			this.lblSec = new System.Windows.Forms.Label();
			this.lblToken = new System.Windows.Forms.Label();
			this.cboRepository = new System.Windows.Forms.ComboBox();
			this.lblURL = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblRepository = new System.Windows.Forms.Label();
			this.btnStop = new System.Windows.Forms.Button();
			this.lblEndDate = new System.Windows.Forms.Label();
			this.lblStartDate = new System.Windows.Forms.Label();
			this.btnSearch = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.nupWait)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblUid
			// 
			this.lblUid.Location = new System.Drawing.Point(0, 32);
			this.lblUid.Name = "lblUid";
			this.lblUid.Size = new System.Drawing.Size(60, 16);
			this.lblUid.TabIndex = 0;
			this.lblUid.Text = "Login ID";
			this.lblUid.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lblPassword
			// 
			this.lblPassword.Location = new System.Drawing.Point(232, 32);
			this.lblPassword.Name = "lblPassword";
			this.lblPassword.Size = new System.Drawing.Size(56, 16);
			this.lblPassword.TabIndex = 1;
			this.lblPassword.Text = "Password";
			this.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cboUid
			// 
			this.cboUid.Items.AddRange(new object[] {
														""});
			this.cboUid.Location = new System.Drawing.Point(64, 28);
			this.cboUid.Name = "cboUid";
			this.cboUid.Size = new System.Drawing.Size(164, 21);
			this.cboUid.TabIndex = 2;
			this.cboUid.Text = "login0.company1.zantaz.com";
			this.ttpGUITest.SetToolTip(this.cboUid, "trim leading and ending space");
			// 
			// txtPassword
			// 
			this.txtPassword.Location = new System.Drawing.Point(288, 28);
			this.txtPassword.Name = "txtPassword";
			this.txtPassword.PasswordChar = '*';
			this.txtPassword.Size = new System.Drawing.Size(96, 20);
			this.txtPassword.TabIndex = 3;
			this.txtPassword.Text = "";
			this.ttpGUITest.SetToolTip(this.txtPassword, "login password");
			this.txtPassword.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPassword_KeyPress);
			// 
			// lnkGuidFile
			// 
			this.lnkGuidFile.Enabled = false;
			this.lnkGuidFile.Location = new System.Drawing.Point(4, 16);
			this.lnkGuidFile.Name = "lnkGuidFile";
			this.lnkGuidFile.Size = new System.Drawing.Size(56, 16);
			this.lnkGuidFile.TabIndex = 4;
			this.lnkGuidFile.TabStop = true;
			this.lnkGuidFile.Text = "GUID File";
			this.lnkGuidFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.ttpGUITest.SetToolTip(this.lnkGuidFile, "Load the GUID file");
			this.lnkGuidFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkGuidFile_LinkClicked);
			// 
			// txtGuidFile
			// 
			this.txtGuidFile.Enabled = false;
			this.txtGuidFile.Location = new System.Drawing.Point(64, 12);
			this.txtGuidFile.Name = "txtGuidFile";
			this.txtGuidFile.Size = new System.Drawing.Size(220, 20);
			this.txtGuidFile.TabIndex = 5;
			this.txtGuidFile.Text = "";
			this.ttpGUITest.SetToolTip(this.txtGuidFile, "GUID file location");
			// 
			// btnLogin
			// 
			this.btnLogin.Location = new System.Drawing.Point(312, 4);
			this.btnLogin.Name = "btnLogin";
			this.btnLogin.Size = new System.Drawing.Size(60, 20);
			this.btnLogin.TabIndex = 6;
			this.btnLogin.Text = "Login";
			this.ttpGUITest.SetToolTip(this.btnLogin, "Login to the DS");
			this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
			// 
			// rtbDisplay
			// 
			this.rtbDisplay.Location = new System.Drawing.Point(4, 240);
			this.rtbDisplay.Name = "rtbDisplay";
			this.rtbDisplay.Size = new System.Drawing.Size(376, 180);
			this.rtbDisplay.TabIndex = 7;
			this.rtbDisplay.Text = "";
			// 
			// cboURL
			// 
			this.cboURL.Items.AddRange(new object[] {
														""});
			this.cboURL.Location = new System.Drawing.Point(64, 4);
			this.cboURL.Name = "cboURL";
			this.cboURL.Size = new System.Drawing.Size(220, 21);
			this.cboURL.TabIndex = 10;
			this.cboURL.Text = "http://10.1.41.243";
			this.ttpGUITest.SetToolTip(this.cboURL, "Need input validation");
			// 
			// txtToken
			// 
			this.txtToken.Location = new System.Drawing.Point(64, 84);
			this.txtToken.Name = "txtToken";
			this.txtToken.Size = new System.Drawing.Size(220, 20);
			this.txtToken.TabIndex = 12;
			this.txtToken.Text = "";
			this.ttpGUITest.SetToolTip(this.txtToken, "In Subject Field - search pattern");
			// 
			// dtpStartDate
			// 
			this.dtpStartDate.CustomFormat = "MM/dd/yyyy hh:mm tt";
			this.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpStartDate.Location = new System.Drawing.Point(64, 108);
			this.dtpStartDate.Name = "dtpStartDate";
			this.dtpStartDate.Size = new System.Drawing.Size(220, 20);
			this.dtpStartDate.TabIndex = 15;
			this.ttpGUITest.SetToolTip(this.dtpStartDate, "12/31/9998 - 1/1/1753");
			this.dtpStartDate.Value = new System.DateTime(1970, 9, 7, 9, 25, 0, 0);
			// 
			// dtpEndDate
			// 
			this.dtpEndDate.CustomFormat = "MM/dd/yyyy hh:mm tt";
			this.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpEndDate.Location = new System.Drawing.Point(64, 132);
			this.dtpEndDate.Name = "dtpEndDate";
			this.dtpEndDate.Size = new System.Drawing.Size(220, 20);
			this.dtpEndDate.TabIndex = 16;
			this.ttpGUITest.SetToolTip(this.dtpEndDate, "12/31/9998 - 1/1/1753");
			this.dtpEndDate.Value = new System.DateTime(2020, 9, 7, 21, 25, 0, 0);
			// 
			// rdoToken
			// 
			this.rdoToken.Checked = true;
			this.rdoToken.Location = new System.Drawing.Point(288, 88);
			this.rdoToken.Name = "rdoToken";
			this.rdoToken.Size = new System.Drawing.Size(16, 16);
			this.rdoToken.TabIndex = 14;
			this.rdoToken.TabStop = true;
			this.rdoToken.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ttpGUITest.SetToolTip(this.rdoToken, "Search in subject field");
			this.rdoToken.Click += new System.EventHandler(this.rdoToken_Click);
			// 
			// rdoGUIDFile
			// 
			this.rdoGUIDFile.Location = new System.Drawing.Point(288, 16);
			this.rdoGUIDFile.Name = "rdoGUIDFile";
			this.rdoGUIDFile.Size = new System.Drawing.Size(16, 16);
			this.rdoGUIDFile.TabIndex = 13;
			this.rdoGUIDFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ttpGUITest.SetToolTip(this.rdoGUIDFile, "Search in subject field");
			this.rdoGUIDFile.Click += new System.EventHandler(this.rdoGUIDFile_Click);
			// 
			// txtQueryFile
			// 
			this.txtQueryFile.Enabled = false;
			this.txtQueryFile.Location = new System.Drawing.Point(64, 36);
			this.txtQueryFile.Name = "txtQueryFile";
			this.txtQueryFile.Size = new System.Drawing.Size(220, 20);
			this.txtQueryFile.TabIndex = 20;
			this.txtQueryFile.Text = "";
			this.ttpGUITest.SetToolTip(this.txtQueryFile, "Query file location");
			// 
			// rdoQueryFile
			// 
			this.rdoQueryFile.Location = new System.Drawing.Point(288, 40);
			this.rdoQueryFile.Name = "rdoQueryFile";
			this.rdoQueryFile.Size = new System.Drawing.Size(16, 16);
			this.rdoQueryFile.TabIndex = 19;
			this.rdoQueryFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ttpGUITest.SetToolTip(this.rdoQueryFile, "Search in Sender field");
			this.rdoQueryFile.Click += new System.EventHandler(this.rdoQueryFile_Click);
			// 
			// txtSubject
			// 
			this.txtSubject.Enabled = false;
			this.txtSubject.Location = new System.Drawing.Point(64, 60);
			this.txtSubject.Name = "txtSubject";
			this.txtSubject.Size = new System.Drawing.Size(220, 20);
			this.txtSubject.TabIndex = 22;
			this.txtSubject.Text = "";
			this.ttpGUITest.SetToolTip(this.txtSubject, "Option: Use subject criteria to limit the search.");
			// 
			// chkOption
			// 
			this.chkOption.Enabled = false;
			this.chkOption.Location = new System.Drawing.Point(288, 64);
			this.chkOption.Name = "chkOption";
			this.chkOption.Size = new System.Drawing.Size(16, 16);
			this.chkOption.TabIndex = 23;
			this.ttpGUITest.SetToolTip(this.chkOption, "Subject field option");
			this.chkOption.Click += new System.EventHandler(this.chkOption_Click);
			// 
			// lnkQueryFile
			// 
			this.lnkQueryFile.Enabled = false;
			this.lnkQueryFile.Location = new System.Drawing.Point(4, 40);
			this.lnkQueryFile.Name = "lnkQueryFile";
			this.lnkQueryFile.Size = new System.Drawing.Size(56, 16);
			this.lnkQueryFile.TabIndex = 24;
			this.lnkQueryFile.TabStop = true;
			this.lnkQueryFile.Text = "Query File";
			this.lnkQueryFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.ttpGUITest.SetToolTip(this.lnkQueryFile, "Load the GUID file");
			this.lnkQueryFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkQueryFile_LinkClicked);
			// 
			// nupWait
			// 
			this.nupWait.Location = new System.Drawing.Point(312, 84);
			this.nupWait.Maximum = new System.Decimal(new int[] {
																	10,
																	0,
																	0,
																	0});
			this.nupWait.Name = "nupWait";
			this.nupWait.Size = new System.Drawing.Size(40, 20);
			this.nupWait.TabIndex = 25;
			this.nupWait.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.ttpGUITest.SetToolTip(this.nupWait, "Rang: 0 .. 10");
			this.nupWait.Value = new System.Decimal(new int[] {
																  2,
																  0,
																  0,
																  0});
			// 
			// lblSec
			// 
			this.lblSec.Location = new System.Drawing.Point(352, 88);
			this.lblSec.Name = "lblSec";
			this.lblSec.Size = new System.Drawing.Size(28, 12);
			this.lblSec.TabIndex = 26;
			this.lblSec.Text = "Sec.";
			this.ttpGUITest.SetToolTip(this.lblSec, "Wait in Second");
			// 
			// lblToken
			// 
			this.lblToken.Location = new System.Drawing.Point(20, 88);
			this.lblToken.Name = "lblToken";
			this.lblToken.Size = new System.Drawing.Size(40, 16);
			this.lblToken.TabIndex = 11;
			this.lblToken.Text = "Token";
			this.lblToken.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.ttpGUITest.SetToolTip(this.lblToken, "GUID ONLY");
			// 
			// cboRepository
			// 
			this.cboRepository.Items.AddRange(new object[] {
															   ""});
			this.cboRepository.Location = new System.Drawing.Point(64, 51);
			this.cboRepository.Name = "cboRepository";
			this.cboRepository.Size = new System.Drawing.Size(220, 21);
			this.cboRepository.TabIndex = 14;
			this.cboRepository.Text = "testdomain1.marketingdirector3.repository";
			this.ttpGUITest.SetToolTip(this.cboRepository, "Respository name");
			// 
			// lblURL
			// 
			this.lblURL.Location = new System.Drawing.Point(0, 8);
			this.lblURL.Name = "lblURL";
			this.lblURL.Size = new System.Drawing.Size(60, 16);
			this.lblURL.TabIndex = 9;
			this.lblURL.Text = "URL";
			this.lblURL.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnStop);
			this.groupBox1.Controls.Add(this.lblSec);
			this.groupBox1.Controls.Add(this.lnkQueryFile);
			this.groupBox1.Controls.Add(this.chkOption);
			this.groupBox1.Controls.Add(this.txtSubject);
			this.groupBox1.Controls.Add(this.txtQueryFile);
			this.groupBox1.Controls.Add(this.rdoQueryFile);
			this.groupBox1.Controls.Add(this.lblEndDate);
			this.groupBox1.Controls.Add(this.lblStartDate);
			this.groupBox1.Controls.Add(this.dtpEndDate);
			this.groupBox1.Controls.Add(this.dtpStartDate);
			this.groupBox1.Controls.Add(this.rdoToken);
			this.groupBox1.Controls.Add(this.rdoGUIDFile);
			this.groupBox1.Controls.Add(this.txtToken);
			this.groupBox1.Controls.Add(this.lnkGuidFile);
			this.groupBox1.Controls.Add(this.lblToken);
			this.groupBox1.Controls.Add(this.txtGuidFile);
			this.groupBox1.Controls.Add(this.btnSearch);
			this.groupBox1.Controls.Add(this.nupWait);
			this.groupBox1.Location = new System.Drawing.Point(0, 76);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(384, 160);
			this.groupBox1.TabIndex = 13;
			this.groupBox1.TabStop = false;
			// 
			// lblRepository
			// 
			this.lblRepository.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.lblRepository.Location = new System.Drawing.Point(0, 56);
			this.lblRepository.Name = "lblRepository";
			this.lblRepository.Size = new System.Drawing.Size(60, 16);
			this.lblRepository.TabIndex = 28;
			this.lblRepository.Text = "Repository";
			this.lblRepository.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// btnStop
			// 
			this.btnStop.Enabled = false;
			this.btnStop.Location = new System.Drawing.Point(340, 132);
			this.btnStop.Name = "btnStop";
			this.btnStop.Size = new System.Drawing.Size(40, 20);
			this.btnStop.TabIndex = 27;
			this.btnStop.Text = "Abort";
			this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
			// 
			// lblEndDate
			// 
			this.lblEndDate.Location = new System.Drawing.Point(4, 136);
			this.lblEndDate.Name = "lblEndDate";
			this.lblEndDate.Size = new System.Drawing.Size(56, 16);
			this.lblEndDate.TabIndex = 18;
			this.lblEndDate.Text = "End Date";
			this.lblEndDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lblStartDate
			// 
			this.lblStartDate.Location = new System.Drawing.Point(4, 112);
			this.lblStartDate.Name = "lblStartDate";
			this.lblStartDate.Size = new System.Drawing.Size(56, 16);
			this.lblStartDate.TabIndex = 17;
			this.lblStartDate.Text = "Start Date";
			this.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// btnSearch
			// 
			this.btnSearch.Enabled = false;
			this.btnSearch.Location = new System.Drawing.Point(288, 132);
			this.btnSearch.Name = "btnSearch";
			this.btnSearch.Size = new System.Drawing.Size(48, 20);
			this.btnSearch.TabIndex = 14;
			this.btnSearch.Text = "Search";
			this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
			// 
			// GUITestPage
			// 
			this.Controls.Add(this.cboRepository);
			this.Controls.Add(this.cboURL);
			this.Controls.Add(this.lblURL);
			this.Controls.Add(this.rtbDisplay);
			this.Controls.Add(this.btnLogin);
			this.Controls.Add(this.txtPassword);
			this.Controls.Add(this.cboUid);
			this.Controls.Add(this.lblPassword);
			this.Controls.Add(this.lblUid);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblRepository);
			this.Name = "GUITestPage";
			this.Size = new System.Drawing.Size(388, 424);
			((System.ComponentModel.ISupportInitialize)(this.nupWait)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Take the URL + login id + password to logging into digital safe.
		/// Format:https://zantaz.digitalsafe.net/zantaz/DigitalSafe/LoginServlet?loginName=kcheung.zantaz.com&password=xxxxxx&JavaScriptDetector=js_enabled
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnLogin_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			if( urlData.getSessionKey() != "" ) 
				HandleLogoff();

			if ( HandleLoginPage() )
				if( HandleSearchPage() )
				{
					btnSearch.Enabled = true;
					m_login = true;
				}// end of if - HandleSearchPage

			this.Cursor = Cursors.Default;
		}//end of btnLogin_Click

		/// <summary>
		/// Get the session id and document class id
		/// </summary>
		/// <returns>1 - login OK, 0 - login fail</returns>
		private bool HandleLoginPage()
		{
			bool rv = true; // assume everything is OK

			cboURL.Text = cboURL.Text.Trim( ' ' ); //trim leading and ending space
			urlData.setBaseURL( cboURL.Text );

			// https://zantaz.digitalsafe.net/zantaz/DigitalSafe/LoginServlet
			string strURL  = cboURL.Text + "/zantaz/DigitalSafe/LoginServlet";
			// loginName=kcheung.zantaz.com&password=xxxxxx&JavaScriptDetector=js_enabled
			string strParam = "loginName=" + cboUid.Text + "&password=" + txtPassword.Text + "&JavaScriptDetector=js_enabled";

			WebResponse result = null;
			try 
			{
				WebRequest req = WebRequest.Create(strURL);
				req.Method = "POST";
				req.ContentType = "application/x-www-form-urlencoded";

				Char[] reserved = {'?', '=', '&'};
				byte[] SomeBytes = null;
				StringBuilder encodedURL = new StringBuilder(); // store encoded url
				if( strParam != null ) // strParam always something (save the if-else for future reference)
				{
					int i = 0;
					int j;
					while( i < strParam.Length )
					{
						j = strParam.IndexOfAny(reserved, i);
						if( j == -1 )
						{
							encodedURL.Append( HttpUtility.UrlEncode(strParam.Substring(i, strParam.Length-i)) );
							break;
						}
						encodedURL.Append(HttpUtility.UrlEncode(strParam.Substring(i, j-i)));
						encodedURL.Append(strParam.Substring(j,1));
						i = j + 1;
					}//end of while

					// For non-ascii language
					SomeBytes = Encoding.UTF8.GetBytes(encodedURL.ToString());
					req.ContentLength = SomeBytes.Length;
					Stream newStream = req.GetRequestStream();
					newStream.Write(SomeBytes, 0, SomeBytes.Length);
					newStream.Close();
				}//end of if 
				else 
				{
					req.ContentLength = 0;
				}

				result = req.GetResponse();
				Stream ReceiveStream = result.GetResponseStream();
				Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
				StreamReader sr = new StreamReader( ReceiveStream, encode );
				Trace.WriteLine("\r\nResponse stream received");

				string line = "";
				while( (line = sr.ReadLine()) != null )
				{
					Debug.WriteLine( line );
					urlData.ExtractDocumentClassList( line );
					urlData.ExtractSessionKey( line );
				}//end of while

				if( (urlData.getDocumentClassList() == "") || (urlData.getSessionKey() == "") )
				{
					rv = false; // fail
					rtbDisplay.Text = "Login fail";
				}
					
				rtbDisplay.Text = "Doc class list = " + urlData.getDocumentClassList() + "\n"
					+ "Session Key = " + urlData.getSessionKey();
			} //end of try
			catch(Exception ex) 
			{
				rv = false; // login fail
				Trace.WriteLine( ex.ToString());
				rtbDisplay.Text = ex.Message.ToString();
			} 
			finally 
			{
				if( result != null ) 
				{
					result.Close();
				}
			}//end of finally
			
			return( rv );
		}//end of HandleLoginPage

		/// <summary>
		/// Get the repository id, which embedded in search page.
		/// </summary>
		/// <returns>1 -  OK, 0 -  fail</returns>
		private bool HandleSearchPage()
		{
			Trace.WriteLine( "GUITestPage.cs - HandleSearchPage" );

			bool rv = true;

			string strURL = urlData.BuildSearchFormURL();
			WebResponse result = null;
			try 
			{
				WebRequest req = WebRequest.Create(strURL);
				result = req.GetResponse();
				Stream ReceiveStream = result.GetResponseStream();
				Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
				StreamReader sr = new StreamReader( ReceiveStream, encode );
				Trace.WriteLine("\r\nResponse stream received");
				string line = "";
				while( (line = sr.ReadLine()) != null )
				{
					Debug.WriteLine( line );
/*
					rtbDisplay.Text = "Doc class list = " + urlData.getDocumentClassList() + "\n"
						+ "Session Key = " + urlData.getSessionKey();
*/

					urlData.ExtractRepositoryID( line, cboRepository.Text );
					rtbDisplay.Text = "Doc class list = " + urlData.getDocumentClassList() + "\n"
						+ "Session Key = " + urlData.getSessionKey() + "\n"
						+ "Repository ID = " + urlData.getRepositoryID();

				}//end of while
			}//end of try 
			catch(WebException ex) 
			{
				rv = false;
				Trace.WriteLine( ex.ToString() );
				rtbDisplay.Text = ex.Message.ToString();
			}//end of catch
			finally 
			{
				if( result != null ) // Is always != null ??
				{
					result.Close();
				}
			}//end of finally

			return( rv );
		}//end of HandleSearchPage

		/// <summary>
		/// Search the HTML
		/// 1) Search input string in HTML   OR
		/// 2) Search GUID from file in HTML OR
		/// 3) Search the pre-defined criteria file.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSearch_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "GUITestPage.cs - btnSearch_Click" );			

			searchThread = new Thread( new ThreadStart(this.ThdWrap_btnSearch) );
			searchThread.Name = "searchThread";
			searchThread.Start();

            commObj.LogToFile("Thread.log", "++ searchThread Start ++");
		}//end of btnSearch_Click

		private void rdoGUIDFile_Click(object sender, System.EventArgs e)
		{
			Debug.WriteLine( "GUITestPage.cs - rdoGUIDFile_Click" );
			txtGuidFile.Enabled  = true;
			lnkGuidFile.Enabled  = true;

			txtQueryFile.Enabled = false;
			lnkQueryFile.Enabled = false;
			chkOption.Checked    = false;			
			chkOption.Enabled	 = false;
			txtSubject.Text		 = "";
			txtSubject.Enabled	 = false;
			
			txtToken.Enabled     = false;
			lblToken.Enabled     = false;
		}//end of rdoGUIDFile_Click

		private void rdoToken_Click(object sender, System.EventArgs e)
		{
			Debug.WriteLine( "GUITestPage.cs - rdoToken_Click" );
			txtGuidFile.Enabled  = false;
			lnkGuidFile.Enabled  = false;

			txtQueryFile.Enabled = false;
			lnkQueryFile.Enabled = false;
			chkOption.Checked    = false;			
			chkOption.Enabled	 = false;
			txtSubject.Text		 = "";
			txtSubject.Enabled	 = false;

			txtToken.Enabled     = true;
			lblToken.Enabled     = true;
		}//end of rdoToken_Click

		private void rdoQueryFile_Click(object sender, System.EventArgs e)
		{
			Debug.WriteLine( "GUITestPage.cs - rdoQueryFile_Click" );
			txtGuidFile.Enabled  = false;
			lnkGuidFile.Enabled  = false;

			txtQueryFile.Enabled = true;
			lnkQueryFile.Enabled = true;
			chkOption.Enabled    = true;

			txtToken.Enabled     = false;
			lblToken.Enabled     = false;		
		}// end of rdoQueryFile_Click

		/// <summary>
		/// Search the input token. It is the subject field.
		/// </summary>
		/// <param name="strToken">input string</param>
		/// <returns>true - found, false - NOT Found</returns>
		public bool doTokenSearch(string strToken)
		{
			Trace.WriteLine("GUITestPage.cs - doTokenSearch");

            strToken = strToken.Trim( ' ' ); // trim leading and ending space
            if( strToken == "" )
                return false; // break

			// assume token search always search in subject field
			urlData.setSubject( strToken ); // update the current urlDataObj

			string fileName = "found";
			string strMsgURL = "";
            string strURL  = urlData.BuildSearchURL();
#if(DEBUG)
            commObj.LogToFile( strURL );
			commObj.LogToFile( "Input search String = " + strToken ); 
#endif            
			WebResponse result = null;
			try 
			{
				doWebRequest( strURL );
				string line = "";
				int idx = 0;
				// OK. After the query search page, get the result page
				strURL = urlData.BuildResultPageURL();
				WebRequest req = WebRequest.Create(strURL);
				result = req.GetResponse();
				Stream ReceiveStream = result.GetResponseStream();
				Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
				StreamReader sr = new StreamReader( ReceiveStream, encode );
				while( (line = sr.ReadLine()) != null )
				{
					Debug.WriteLine( line );

 					if( (idx = line.IndexOf(strToken)) != -1 ) // -1 not found
					{
						commObj.LogToFile( "FOUND " + m_found.ToString() + " - " + line );
						commObj.WriteLineByLine( "guidFOUND.txt", strToken );
						strMsgURL = urlData.BuildMsgDisplayURL( line );
						HandleDisplayMsg( fileName + m_found.ToString() + ".html", strMsgURL );
						m_found++;
					}//end of if - found the string					 
				}//end of while
			}//end of try 
			catch(WebException ex) 
			{
				Trace.WriteLine( ex.ToString() );
				rtbDisplay.Text = ex.Message.ToString();
			}//end of catch
			finally 
			{
				if( result != null ) // Is always != null ??
				{
					result.Close();
				}
			}//end of finally
			return true;
		}//end of doTokenSearch

		/// <summary>
		/// Fill whatever value we got from UI
		/// </summary>
		public void FillURLDataObj()
		{
			txtToken.Text = txtToken.Text.Trim( ' ' ); // trim leading and ending space
			urlData.setSubject( txtToken.Text ); // save the token: Search the subject field only

//			cboURL.Text = cboURL.Text.Trim( ' ' ); //trim leading and ending space
//			urlData.setBaseURL( cboURL.Text );

			urlData.setStartYear( dtpStartDate.Value.Year.ToString() );
			urlData.setStartMonth( dtpStartDate.Value.Month.ToString() );
			urlData.setStartDay( dtpStartDate.Value.Day.ToString() );
			urlData.setStartHour( dtpStartDate.Value.Hour.ToString() );
			urlData.setStartMinute( dtpStartDate.Value.Minute.ToString() );

			urlData.setEndYear( dtpEndDate.Value.Year.ToString() );
			urlData.setEndMonth( dtpEndDate.Value.Month.ToString() );
			urlData.setEndDay( dtpEndDate.Value.Day.ToString() );
			urlData.setEndHour( dtpEndDate.Value.Hour.ToString() );
			urlData.setEndMinute( dtpEndDate.Value.Minute.ToString() );			
		}//end of FillURLDataObj

		/// <summary>
		/// Save found message
		/// </summary>
		/// <param name="fn">file name for each found message</param>
		/// <param name="strURL">the message display url</param>
		public void HandleDisplayMsg( string fn, string strURL )
		{
			Trace.WriteLine( "GUITestPage.cs - HandleDisplayMsg" );

			WebResponse result = null;
			try 
			{
				WebRequest req = WebRequest.Create(strURL);
				result = req.GetResponse();
				Stream ReceiveStream = result.GetResponseStream();
				Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
				StreamReader sr = new StreamReader( ReceiveStream, encode );
				Trace.WriteLine("\r\nMessage stream received");
				string line = "";
				while( (line = sr.ReadLine()) != null )
				{
					Debug.WriteLine( line );
					commObj.WriteLineByLine( fn, line );
				}//end of while
			}//end of try 
			catch(WebException ex) 
			{
				Trace.WriteLine( ex.ToString() );
				rtbDisplay.Text = ex.Message.ToString();
			}//end of catch
			finally 
			{
				if( result != null ) // Is always != null ??
				{
					result.Close();
				}
			}//end of finally		
		}//end of HandleDisplayMsg

		/// <summary>
		/// Get the GUID file name
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void lnkGuidFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtGuidFile.Text = ofDlg.FileName;
			}//end of if		
		}//end of lnkGuidFile_LinkClicked

		private void lnkQueryFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.RestoreDirectory = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtQueryFile.Text = ofDlg.FileName;
			}//end of if				
		}// end of lnkQueryFile_LinkClicked

		/// <summary>
		/// Read from GUID file to get the unique GUID, then search on the guid
		/// Search in subject field
		/// </summary>
		/// <param name="fn">absolute file name and path for the guid file</param>
		public void doGuidFileSearch(string fn)
		{
			string line = "";
			if( fn != "" )
			{
				StreamReader sr = new StreamReader( fn );
				while( (line = sr.ReadLine()) != null )
				{						
					Debug.WriteLine( "GUID: " + line );
					doTokenSearch( line );
				}//end of while
				sr.Close();
			}//end of if - open file name
			else
			{	// fn == "" - ERROR				
				rtbDisplay.Text += "\nCheck the file name";
			}//end of else
		}// end of doGuidFileSearch

		/// <summary>
		/// Enable state of txtSubject is depend on the option check box
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkOption_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "GUITestPage.cs - chkOption_Click" );
			if( chkOption.Checked )
			{
				txtSubject.Enabled = true;
			}
			else
			{
				txtSubject.Text = "";
				txtSubject.Enabled = false;
			}
		}// end of chkOption_Click

		/// <summary>
		/// Replace +, &, |, (, ), ! to its corresponding HEX.
		/// Since space will be replaced by '+', replace + sign first.
		/// </summary>
		/// <param name="strQuery"></param>
		/// <returns></returns>
		public string RebuildTokenString(string strQuery)
		{
			strQuery = strQuery.Replace("+", "%2B");
			strQuery = strQuery.Replace("&", "%26");
			strQuery = strQuery.Replace("|", "%7C");
			strQuery = strQuery.Replace("(", "%28");
			strQuery = strQuery.Replace(")", "%29");
			strQuery = strQuery.Replace("!", "%21");
			strQuery = strQuery.Replace(' ', '+' );
			return( strQuery );
		}//end of RebuildTokenString

		/// <summary>
		/// Check the HTML line contain the expected result (guid) or not.
		/// Only the first return page (first 20 result) will check. 
		/// Assume the pre-defined result (expected result) is correct.
		/// Assume the archive function is OK; ie: only check the index function here.
		/// </summary>
		/// <param name="line">HTML line that long enough to check</param>
		/// <param name="strArray">array of expected result</param>
		/// <param name="total">total found in this HTML page</param>
		/// <returns>Number of message found</returns>
		public int CheckExpectedResult(string line, string[] strArray, int total)
		{
			int num = strArray.Length;
			for( int i = 1; i < num; i++ )// don't count the first one, start from 2nd		
			{				
				if( (line.IndexOf(strArray[i].ToString())) != -1 ) // found
				{
					commObj.WriteLineByLine( "Result" + m_found.ToString() + ".txt", strArray[i].ToString() );
					total++; 
				}//end of if - found
			}//end of for
			return( total );
		}// end of CheckExpectedResult

		/// <summary>
		/// Read the query criteria and perform search on the sender field (From field)
		/// Have option to include the query in subject box.
		/// </summary>
		/// <param name="fn"></param>
		public void doQueryFileSearch(string fn)
		{
			if( fn == "" )
			{
				rtbDisplay.Text += "\nCheck the file name";
				return;
			}

			if( chkOption.Enabled && chkOption.Checked ) // take the option subject field
			{
				string strSubject = txtSubject.Text.Trim(' ');
				if( strSubject != "" )
				{
					urlData.setSubject( strSubject );
				}// end of if
			}//end of if

			string line = "";
			string fileName = "TC"; // file name prefix for HTML page
			string strURL  = "";
			StreamReader fsr = new StreamReader( fn );
			while( (line = fsr.ReadLine()) != null )
			{
				Debug.WriteLine( "TC: " + line );
				string [] splitStr = line.Split( new Char [] {';'} );
				for( int k = 0; k < splitStr.Length; k++ ) // trim leading and ending space
					splitStr[k] = splitStr[k].Trim(' ');

				line = RebuildTokenString( splitStr[0].ToString() );

				////////////////////////////////////////////////////
				urlData.setFrom( line ); // set the from field
				///////////////////////////////////////////////////
				strURL = urlData.BuildSearchURL();
#if(DEBUG)
				commObj.LogToFile( "strURL: " + strURL );
#endif          
				// create the expected result file with the query criteria string
				commObj.WriteLineByLine( "Result" + m_found.ToString() + ".txt", splitStr[0].ToString() );
				int expectedResult = splitStr.Length - 1; // don't count the first element, which is query criteria string
				WebResponse result = null;
				try 
				{
					doWebRequest( strURL );
					// OK. After the query search page, get the result page
					int numFound = 0; // the returned number of result per HTML page
					strURL = urlData.BuildResultPageURL();
					WebRequest req = WebRequest.Create(strURL);

					Thread.Sleep( (int)nupWait.Value * 1000 ); // delay between 2 pages

					result = req.GetResponse();
					Stream ReceiveStream = result.GetResponseStream();
					Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
					StreamReader sr = new StreamReader( ReceiveStream, encode );
					while( (line = sr.ReadLine()) != null )
					{
						Debug.WriteLine( line );
						commObj.WriteLineByLine( fileName + m_found.ToString() + ".html", line );

						// If the HTML line is too short, no need to search.
						// CheckExpectedResult will check the number of return result from HTML page.
						if( CHECK_SIZE < line.Length )
							numFound = CheckExpectedResult( line, splitStr, numFound ); // and write the result to result file
					}//end of while
					// record the num of result in the end of Result file
					commObj.WriteLineByLine( "Result" + m_found.ToString() + ".txt", numFound.ToString() );

					WriteToCheckLog( expectedResult, numFound, splitStr[0].ToString() );

				}//end of try 
				catch(WebException ex) 
				{
					Trace.WriteLine( ex.ToString() );
					rtbDisplay.Text = ex.Message.ToString();
				}//end of catch
				finally 
				{
					if( result != null ) // Is always != null ??
					{
						result.Close();
					}
				}//end of finally
				m_found++; // update the file name
			}//end of while
		}// end of doQueryFileSearch

		/// <summary>
		/// Log the result per page
		/// </summary>
		/// <param name="expected"></param>
		/// <param name="found"></param>
		public void WriteToCheckLog(int expected, int found, string queryStr)
		{
			string str = expected + " : " + found + " - Result" + m_found.ToString() + ".txt" + " - " + queryStr;
			if( (found != expected) || (found == 0 ) ) // don't count the 1st element, which is query string
			{
				str = "Check " + str;
				rtbDisplay.Text += "\n++ " + str;
			}//end of if - NOT Match
			commObj.WriteLineByLine( "Check.log", str );		
		}//end of WriteToCheckLog

		/// <summary>
		/// Do a signle web request. Specific for logoff.
		/// </summary>
		/// <param name="url"></param>
		public void doWebRequest(string url)
		{
			Trace.WriteLine( "GUITestPage.cs - doWebRequest" );

			WebResponse result = null;
			try 
			{
				WebRequest req = WebRequest.Create(url);
//save				Thread.Sleep( (int)nupWait.Value * 1000 ); // delay between 2 pages
				result = req.GetResponse();
				Stream ReceiveStream = result.GetResponseStream();
				Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
				StreamReader sr = new StreamReader( ReceiveStream, encode );
				Trace.WriteLine("\r\nMessage stream received");
				string line = "";
				while( (line = sr.ReadLine()) != null )
				{
					Debug.WriteLine( line );
				}//end of while
			}//end of try 
			catch(WebException ex) 
			{
				Trace.WriteLine( ex.ToString() );
				rtbDisplay.Text = ex.Message.ToString();
			}//end of catch
			finally 
			{
				if( result != null ) // Is always != null ??
				{
					result.Close();
				}
			}//end of finally		
		}// end of doWebRequest

		/// <summary>
		/// Perform logoff - when app close
		/// </summary>
		public void HandleLogoff()
		{
			if( urlData.getSessionKey() != "" )
				doWebRequest( urlData.BuildLogoffURL() );
		}

		/// <summary>
		/// ThreadStart delegate has no parameters or return value. Therefore, cannot start a thread 
		/// using a method that takes paramenters or obtain a return value from the method.
		/// In this case, use a wrapper method to wrap it.
 		/// Search the HTML
		/// 1) Search input string in HTML   OR
		/// 2) Search GUID from file in HTML OR
		/// 3) Search the pre-defined criteria file.
		/// Also disable the login & search button until the search finished. No one can login and re-issue search
		/// </summary>
		public void ThdWrap_btnSearch()
		{
			this.Cursor = Cursors.WaitCursor;
			btnLogin.Enabled  = false;
			btnSearch.Enabled = false;
			btnStop.Enabled   = true;

			m_found = 0; // reset file counter			
			FillURLDataObj(); // save GUI value into URLDataObj

			if( rdoGUIDFile.Checked )
			{
				Debug.WriteLine( "\tDo the GUID File part" );
				doGuidFileSearch( txtGuidFile.Text );
			}
			else
				if( rdoToken.Checked )
			{
				Debug.WriteLine( "\tDo the token search part" );
				doTokenSearch( txtToken.Text );
			}
			else			
				if( rdoQueryFile.Checked )
			{
				Debug.WriteLine( "\tDo the Query File search part" );
				doQueryFileSearch( txtQueryFile.Text );
			}//end of else - rdoQueryFile

			rtbDisplay.Text += "\nMessage found = " + m_found.ToString();

			btnLogin.Enabled  = true;
			btnSearch.Enabled = true;
			btnStop.Enabled   = false;
			this.Cursor = Cursors.Default;
		}// end of ThdWrap_btnSearch

		/// <summary>
		/// Do the login when user hit enter key by calling btnLogin_Click method.
		/// ATTEN: Since btnLogin_Click method doesn't use the input parameter, the input parameters here
		/// contain dummy value. If input parameters use in future, it will break here.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void txtPassword_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Trace.WriteLine( "GUITestPage.cs - txtPassword_KeyPress" );
			if( e.KeyChar == 0x0D ) // enter key hit
				btnLogin_Click(this,  e); // both input params are NOT use. If use, will break.
		}//end of txtPassword_KeyPress

		/// <summary>
		/// Kill the search thread when the program exit.
		/// </summary>
		public void KillSearchThread()
		{
			Trace.WriteLine("GUITestPage.cs - KillSearchThread()");
			try
			{
                commObj.LogToFile("Thread.log", "++Kill Thread:" + searchThread.Name );
				searchThread.Abort(); // abort
                searchThread.Join();  // require for ensure the thread kill
			}//end of try 
			catch( ThreadAbortException thdEx )
			{
				Trace.WriteLine( thdEx.Message );
				rtbDisplay.Text += "\nAborting the search thread";
			}//end of catch
		}//end of KillSearchThread()

		/// <summary>
		/// Abort the searching thread... and enable the GUI button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnStop_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine("GUITestPage.cs - btnStop_Click");
			KillSearchThread();

			btnLogin.Enabled  = true; // enable both login button and search button
			btnSearch.Enabled = true;
			btnStop.Enabled   = false;
			this.Cursor = Cursors.Default; // reset cursor
		}

	}
}
