using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Web;
using System.Web.Mail;

namespace QATool
{
	/// <summary>
	/// Summary description for BatchMail.
	/// </summary>
	public class BatchMail : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.ComboBox cboPort;
		private System.Windows.Forms.ComboBox cboSMTP;
		private System.Windows.Forms.Label lblPort;
		private System.Windows.Forms.Label lblSMTP;
		private System.Windows.Forms.RichTextBox richBox;
		private System.Windows.Forms.Label lblSubject;
		private System.Windows.Forms.TextBox txtSubject;
		private System.Windows.Forms.TextBox txtBCC;
		private System.Windows.Forms.TextBox txtCC;
		private System.Windows.Forms.TextBox txtTo;
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
		private System.Windows.Forms.CheckBox chkCC;
		private System.Windows.Forms.CheckBox chkBCC;
		private System.Windows.Forms.CheckBox chkGUID; 
		private QATool.CommObj commObj = new CommObj();

		public BatchMail()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
			ttpBatchMail.SetToolTip( lnkFrom,  "Load the address book" );
			ttpBatchMail.SetToolTip( lnkTo,    "Load the address book" );
			ttpBatchMail.SetToolTip( lnkCC,    "Load the address book" );
			ttpBatchMail.SetToolTip( lnkBCC,   "Load the address book" );
			ttpBatchMail.SetToolTip( lnkFolder,"Specific attachment data location" );
			ttpBatchMail.SetToolTip( lblLoop,  "Repeat count - 1 .. 9999" );
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
			this.btnTest = new System.Windows.Forms.Button();
			this.cboPort = new System.Windows.Forms.ComboBox();
			this.cboSMTP = new System.Windows.Forms.ComboBox();
			this.lblPort = new System.Windows.Forms.Label();
			this.lblSMTP = new System.Windows.Forms.Label();
			this.richBox = new System.Windows.Forms.RichTextBox();
			this.lblSubject = new System.Windows.Forms.Label();
			this.txtSubject = new System.Windows.Forms.TextBox();
			this.txtBCC = new System.Windows.Forms.TextBox();
			this.txtCC = new System.Windows.Forms.TextBox();
			this.txtTo = new System.Windows.Forms.TextBox();
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
			this.chkCC = new System.Windows.Forms.CheckBox();
			this.chkBCC = new System.Windows.Forms.CheckBox();
			this.chkGUID = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize)(this.nudLoop)).BeginInit();
			this.SuspendLayout();
			// 
			// btnTest
			// 
			this.btnTest.Location = new System.Drawing.Point(316, 176);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(64, 21);
			this.btnTest.TabIndex = 35;
			this.btnTest.Text = "Test";
			this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
			// 
			// cboPort
			// 
			this.cboPort.ItemHeight = 13;
			this.cboPort.Location = new System.Drawing.Point(268, 176);
			this.cboPort.Name = "cboPort";
			this.cboPort.Size = new System.Drawing.Size(44, 21);
			this.cboPort.Sorted = true;
			this.cboPort.TabIndex = 34;
			this.cboPort.Text = "25";
			// 
			// cboSMTP
			// 
			this.cboSMTP.ItemHeight = 13;
			this.cboSMTP.Location = new System.Drawing.Point(88, 176);
			this.cboSMTP.Name = "cboSMTP";
			this.cboSMTP.Size = new System.Drawing.Size(132, 21);
			this.cboSMTP.Sorted = true;
			this.cboSMTP.TabIndex = 33;
			this.cboSMTP.Text = "10.1.89.201";
			// 
			// lblPort
			// 
			this.lblPort.Location = new System.Drawing.Point(224, 180);
			this.lblPort.Name = "lblPort";
			this.lblPort.Size = new System.Drawing.Size(38, 16);
			this.lblPort.TabIndex = 32;
			this.lblPort.Text = "Port # ";
			// 
			// lblSMTP
			// 
			this.lblSMTP.Location = new System.Drawing.Point(4, 180);
			this.lblSMTP.Name = "lblSMTP";
			this.lblSMTP.Size = new System.Drawing.Size(72, 16);
			this.lblSMTP.TabIndex = 31;
			this.lblSMTP.Text = "SMTP Server";
			// 
			// richBox
			// 
			this.richBox.Location = new System.Drawing.Point(10, 204);
			this.richBox.Name = "richBox";
			this.richBox.Size = new System.Drawing.Size(372, 164);
			this.richBox.TabIndex = 30;
			this.richBox.Text = "richBox";
			// 
			// lblSubject
			// 
			this.lblSubject.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubject.Location = new System.Drawing.Point(4, 108);
			this.lblSubject.Name = "lblSubject";
			this.lblSubject.Size = new System.Drawing.Size(56, 16);
			this.lblSubject.TabIndex = 29;
			this.lblSubject.Text = "Subject :";
			this.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSubject
			// 
			this.txtSubject.Location = new System.Drawing.Point(68, 104);
			this.txtSubject.Name = "txtSubject";
			this.txtSubject.Size = new System.Drawing.Size(312, 20);
			this.txtSubject.TabIndex = 28;
			this.txtSubject.Text = "txtSubject";
			// 
			// txtBCC
			// 
			this.txtBCC.Enabled = false;
			this.txtBCC.Location = new System.Drawing.Point(68, 80);
			this.txtBCC.Name = "txtBCC";
			this.txtBCC.Size = new System.Drawing.Size(312, 20);
			this.txtBCC.TabIndex = 27;
			this.txtBCC.Text = "";
			// 
			// txtCC
			// 
			this.txtCC.Enabled = false;
			this.txtCC.Location = new System.Drawing.Point(68, 56);
			this.txtCC.Name = "txtCC";
			this.txtCC.Size = new System.Drawing.Size(312, 20);
			this.txtCC.TabIndex = 26;
			this.txtCC.Text = "";
			// 
			// txtTo
			// 
			this.txtTo.Location = new System.Drawing.Point(68, 32);
			this.txtTo.Name = "txtTo";
			this.txtTo.Size = new System.Drawing.Size(312, 20);
			this.txtTo.TabIndex = 25;
			this.txtTo.Text = "login0@company1.zataz.com";
			// 
			// txtFrom
			// 
			this.txtFrom.Location = new System.Drawing.Point(68, 8);
			this.txtFrom.Name = "txtFrom";
			this.txtFrom.Size = new System.Drawing.Size(312, 20);
			this.txtFrom.TabIndex = 24;
			this.txtFrom.Text = "txtFrom";
			// 
			// lnkBCC
			// 
			this.lnkBCC.Enabled = false;
			this.lnkBCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lnkBCC.Location = new System.Drawing.Point(24, 80);
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
			this.lnkCC.Enabled = false;
			this.lnkCC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lnkCC.Location = new System.Drawing.Point(32, 60);
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
			this.lnkTo.Location = new System.Drawing.Point(4, 36);
			this.lnkTo.Name = "lnkTo";
			this.lnkTo.Size = new System.Drawing.Size(56, 16);
			this.lnkTo.TabIndex = 21;
			this.lnkTo.TabStop = true;
			this.lnkTo.Text = "To :";
			this.lnkTo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.lnkTo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkTo_LinkClicked);
			// 
			// lnkFrom
			// 
			this.lnkFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lnkFrom.Location = new System.Drawing.Point(4, 12);
			this.lnkFrom.Name = "lnkFrom";
			this.lnkFrom.Size = new System.Drawing.Size(56, 16);
			this.lnkFrom.TabIndex = 20;
			this.lnkFrom.TabStop = true;
			this.lnkFrom.Text = "From :";
			this.lnkFrom.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.lnkFrom.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFrom_LinkClicked);
			// 
			// lblLoop
			// 
			this.lblLoop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblLoop.Location = new System.Drawing.Point(248, 132);
			this.lblLoop.Name = "lblLoop";
			this.lblLoop.Size = new System.Drawing.Size(56, 16);
			this.lblLoop.TabIndex = 36;
			this.lblLoop.Text = "# of Loop";
			// 
			// nudLoop
			// 
			this.nudLoop.Location = new System.Drawing.Point(316, 128);
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
			this.nudLoop.TabIndex = 39;
			this.nudLoop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.ttpBatchMail.SetToolTip(this.nudLoop, "1 .. 9999");
			this.nudLoop.Value = new System.Decimal(new int[] {
																  1,
																  0,
																  0,
																  0});
			// 
			// chkAttach
			// 
			this.chkAttach.Location = new System.Drawing.Point(4, 132);
			this.chkAttach.Name = "chkAttach";
			this.chkAttach.Size = new System.Drawing.Size(80, 16);
			this.chkAttach.TabIndex = 40;
			this.chkAttach.Text = "Attachment";
			this.chkAttach.CheckedChanged += new System.EventHandler(this.chkAttach_CheckedChanged);
			// 
			// btnSend
			// 
			this.btnSend.Location = new System.Drawing.Point(316, 151);
			this.btnSend.Name = "btnSend";
			this.btnSend.Size = new System.Drawing.Size(64, 21);
			this.btnSend.TabIndex = 41;
			this.btnSend.Text = "Send";
			this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
			// 
			// lnkFolder
			// 
			this.lnkFolder.Enabled = false;
			this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lnkFolder.Location = new System.Drawing.Point(4, 152);
			this.lnkFolder.Name = "lnkFolder";
			this.lnkFolder.Size = new System.Drawing.Size(48, 16);
			this.lnkFolder.TabIndex = 42;
			this.lnkFolder.TabStop = true;
			this.lnkFolder.Text = "Folder :";
			this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.lnkFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkFolder_LinkClicked);
			// 
			// txtFolder
			// 
			this.txtFolder.Enabled = false;
			this.txtFolder.Location = new System.Drawing.Point(88, 152);
			this.txtFolder.Name = "txtFolder";
			this.txtFolder.Size = new System.Drawing.Size(132, 20);
			this.txtFolder.TabIndex = 43;
			this.txtFolder.Text = "txtFolder";
			// 
			// chkCC
			// 
			this.chkCC.Location = new System.Drawing.Point(8, 60);
			this.chkCC.Name = "chkCC";
			this.chkCC.Size = new System.Drawing.Size(16, 16);
			this.chkCC.TabIndex = 44;
			this.chkCC.Text = "CC";
			this.chkCC.CheckedChanged += new System.EventHandler(this.chkCC_CheckedChanged);
			// 
			// chkBCC
			// 
			this.chkBCC.Location = new System.Drawing.Point(8, 84);
			this.chkBCC.Name = "chkBCC";
			this.chkBCC.Size = new System.Drawing.Size(16, 16);
			this.chkBCC.TabIndex = 45;
			this.chkBCC.Text = "BCC";
			this.chkBCC.CheckedChanged += new System.EventHandler(this.chkBCC_CheckedChanged);
			// 
			// chkGUID
			// 
			this.chkGUID.Location = new System.Drawing.Point(88, 132);
			this.chkGUID.Name = "chkGUID";
			this.chkGUID.Size = new System.Drawing.Size(124, 16);
			this.chkGUID.TabIndex = 46;
			this.chkGUID.Text = "Include GUID";
			// 
			// BatchMail
			// 
			this.Controls.Add(this.chkGUID);
			this.Controls.Add(this.chkBCC);
			this.Controls.Add(this.chkCC);
			this.Controls.Add(this.txtFolder);
			this.Controls.Add(this.lnkFolder);
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
			this.Controls.Add(this.txtBCC);
			this.Controls.Add(this.txtCC);
			this.Controls.Add(this.txtTo);
			this.Controls.Add(this.txtFrom);
			this.Controls.Add(this.lnkBCC);
			this.Controls.Add(this.lnkCC);
			this.Controls.Add(this.lnkTo);
			this.Controls.Add(this.lnkFrom);
			this.Name = "BatchMail";
			this.Size = new System.Drawing.Size(388, 376);
			((System.ComponentModel.ISupportInitialize)(this.nudLoop)).EndInit();
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
			Trace.WriteLine( "BatchMail.cs - lnkFrom_LinkClicked" );

			OpenFileDialog ofDlg = new OpenFileDialog();
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtFrom.Text = ofDlg.FileName;
			}//end of if		
		}// end of lnkFrom_LinkClicked

		private void lnkTo_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatchMail.cs - lnkTo_LinkClicked" );
/** save for future
			OpenFileDialog ofDlg = new OpenFileDialog();
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtTo.Text = ofDlg.FileName;
			}//end of if		
** save for future **/
		}// end of lnkTo_LinkClicked

		private void lnkCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatchMail.cs - lnkTo_LinkClicked" );
/** save for future
			OpenFileDialog ofDlg = new OpenFileDialog();
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtCC.Text = ofDlg.FileName;
			}//end of if		
** save for future **/			
		}// end of lnkCC_LinkClicked

		private void lnkBCC_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatchMail.cs - lnkTo_LinkClicked" );
/** save for future
			OpenFileDialog ofDlg = new OpenFileDialog();
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				txtBCC.Text = ofDlg.FileName;
			}//end of if		
** save for future **/			
		}// end of lnkBCC_LinkClicked

		private void btnTest_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine( "BatchMail.cs - lnkTest_LinkClicked" );
			this.Cursor = Cursors.WaitCursor;
			try
			{
				// If SmtpServer is not set, local SMTP server is used
				SmtpMail.SmtpServer = cboSMTP.Text;

				// Test the connection - if smtp server down, exit.
				System.Net.Sockets.TcpClient tcpClient = new TcpClient();
				tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboPort.Text) );

				tcpClient.Close();
				richBox.Text = "Connection is ok....";
			}//end of try
			catch( SocketException ex )
			{
				Debug.WriteLine(ex.Message.ToString());
				MessageBox.Show(ex.Message.ToString(), msgCaption);
			}//end of catch - SocketException		

			this.Cursor = Cursors.Default;		
		}// end of btnTest_Click

		/// <summary>
		/// Include attachment - 5 different files:
		/// test0.pdf, test1.doc, test2.exe, test3.txt, test4.zip
		/// </summary>
		public String getAttachmentName( ref int i )
		{
			Trace.WriteLine("BatchMailForm.cs - getAttachmentName()");
			String sz = "";
			switch( i )
			{
				case 0:
					sz = txtFolder.Text + "\\test0.zip"; i++;
					break;
				case 1:
					sz = txtFolder.Text + "\\test1.pdf"; i++;
					break;
				case 2:
					sz = txtFolder.Text + "\\test2.doc"; i++;
					break;
				case 3:
					sz = txtFolder.Text + "\\test3.xls"; i++;
					break;
				case 4:
					sz = txtFolder.Text + "\\test4.doc"; i = 0; // reset i
					break;
			}//end of switch - i
			return(sz);
		}// end of getAttachmentName

		private void btnSend_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine("BatchMail.cs - btnSend_Click");
			// check number of repeat loop
			// send until loop done
			this.Cursor = Cursors.WaitCursor;
//			String inStr = txtSubject.Text; // save the user input

			try
			{
				System.Net.Sockets.TcpClient tcpClient = new TcpClient();
				tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(cboPort.Text) );
			
				Debug.WriteLine("  +  Reading file - " + txtFrom.Text);

				StreamReader sr = new StreamReader( txtFrom.Text );
				HandleSendMail( sr );
/**
				String strGUID = "";
				String strFrom;
				int    i = 0;
				while( (strFrom = sr.ReadLine()) != null ) // file name from txtFrom field
				{
					richBox.Text += "\r\n- Read line : " + strFrom;
					if( chkGUID.Checked )
					{
						strGUID = System.Guid.NewGuid().ToString();
						txtSubject.Text = inStr + " " + strGUID;
						commObj.LogGUID( "GUID.LOG", strGUID );
					}//end of if - GUID			

					MailMessage mailMsg = new MailMessage();

					mailMsg.From	= strFrom;		// single line from file
					mailMsg.To		= txtTo.Text;	// single line from GUI
					mailMsg.Cc		= txtCC.Text;	// user input, may != To
					mailMsg.Bcc		= txtBCC.Text;	// user input, may != To
					mailMsg.Subject = txtSubject.Text;
					mailMsg.Body	= richBox.Text + "\n" + strFrom + "\n" + DateTime.Now + " - " + txtSubject.Text;

					if( chkAttach.Checked )
					{
						lnkFolder.Enabled = true;
						txtFolder.Enabled = true;
						mailMsg.Attachments.Add( new MailAttachment(getAttachmentName( ref i ), MailEncoding.Base64) );
					}//end of if

					try
					{
						Debug.WriteLine("  +  btnSend_Click - send mail la");
						richBox.Text += "\r\n+ Do the send mail";
						SmtpMail.Send( mailMsg );
					}//end of try
					catch( System.Web.HttpException ex )
					{
						Debug.WriteLine(ex.Message.ToString());
						MessageBox.Show(ex.Message.ToString(), msgCaption);
					}// end of catch

				}//end of while
				richBox.Text += "\r\n+ Batch Mails Sent to " + strFrom;				
				sr.Close();		
**/				
				tcpClient.Close();
			}// end of try
			catch( Exception ex )
			{
				Debug.WriteLine( ex.Message.ToString() );
			}//end of catch - generic exception

			this.Cursor = Cursors.Default;
		}//end of btnSend_Click

		private void chkCC_CheckedChanged(object sender, System.EventArgs e)
		{
			if( chkCC.Checked )
			{
				lnkCC.Enabled = txtCC.Enabled = true;
				txtCC.Text = txtTo.Text; // set CC == To
			}
			else
			{
				lnkCC.Enabled = txtCC.Enabled = false;
				txtCC.Text = ""; 
			}
		}//end of chkCC_CheckedChanged

		private void chkBCC_CheckedChanged(object sender, System.EventArgs e)
		{
			if( chkBCC.Checked )
			{
				lnkBCC.Enabled = txtBCC.Enabled = true;
				txtBCC.Text = txtTo.Text; // set BCC == To
			}
			else
			{
				lnkBCC.Enabled = txtBCC.Enabled = false;
				txtBCC.Text = "";
			}		
		}// end of chkBCC_CheckedChanged

		private void lnkFolder_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			Trace.WriteLine( "BatchMail.cs - lnkTo_LinkClicked" );
			FolderBrowserDialog fbDlg = new FolderBrowserDialog();
			if( fbDlg.ShowDialog() == DialogResult.OK )
			{
				txtFolder.Text = fbDlg.SelectedPath;
			}

		}

		/// <summary>
		/// Constructing the mail based on the GUI setting and then send the mail
		/// Check - inlcude GUID in subject line and body
		/// Log the GUID into a text file for future searching
		/// Check - include attachment 
		/// </summary>
		public void HandleSendMail( StreamReader sr )
		{
			String strGUID = "";
			String strFrom;
			String inStr = txtSubject.Text; // save the user input

			int    i = 0;
			while( (strFrom = sr.ReadLine()) != null ) // file name from txtFrom field
			{
				richBox.Text += "\r\n- Read line : " + strFrom;
				if( chkGUID.Checked )
				{
					strGUID = System.Guid.NewGuid().ToString();
					txtSubject.Text = inStr + " " + strGUID;
					commObj.LogGUID( "GUID.LOG", strGUID );
				}//end of if - GUID			

				MailMessage mailMsg = new MailMessage();

				mailMsg.From	= strFrom;		// single line from file
				mailMsg.To		= txtTo.Text;	// single line from GUI
				mailMsg.Cc		= txtCC.Text;	// user input, may != To
				mailMsg.Bcc		= txtBCC.Text;	// user input, may != To
				mailMsg.Subject = txtSubject.Text;
				mailMsg.Body	= richBox.Text + "\n" + strFrom + "\n" + DateTime.Now + " - " + txtSubject.Text;

				if( chkAttach.Checked )
				{
					lnkFolder.Enabled = true;
					txtFolder.Enabled = true;
					mailMsg.Attachments.Add( new MailAttachment(getAttachmentName( ref i ), MailEncoding.Base64) );
				}//end of if

				try
				{
					Debug.WriteLine("  +  btnSend_Click - send mail la");
					richBox.Text += "\r\n+ Do the send mail";
//					SmtpMail.Send( mailMsg );
				}//end of try
				catch( System.Web.HttpException ex )
				{
					Debug.WriteLine(ex.Message.ToString());
					MessageBox.Show(ex.Message.ToString(), msgCaption);
				}// end of catch

			}//end of while
			richBox.Text += "\r\n+ Batch Mails Sent to " + strFrom;				
			sr.Close();				
		}// end of chkBCC_CheckedChanged
	}
}
