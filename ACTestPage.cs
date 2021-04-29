using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace QATool
{

	/// <summary>
	/// Summary description for ACTestPage.
	/// </summary>
	public class ACTestPage : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.ToolTip ttpACPage;
		private System.Windows.Forms.Button btnPSTCheck;
		private System.Windows.Forms.TextBox txtGuidFileName;
		private System.Windows.Forms.LinkLabel lnkGUIDFile;
		private System.Windows.Forms.LinkLabel lnkPSTFile;
		private System.Windows.Forms.TextBox txtPSTFileName;
		private System.Windows.Forms.LinkLabel lnkNSFFile;
		private System.Windows.Forms.TextBox txtNSFFileName;
		private System.Windows.Forms.Button btnNSFCheck;
		private System.Windows.Forms.CheckBox chkOutFile;
		private System.Windows.Forms.TextBox txtLogFileName;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.TextBox txtOutput;
		private System.Windows.Forms.ListBox lsbDisplay;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtPassword;
		private System.Windows.Forms.LinkLabel lnkProfile;
		private System.Windows.Forms.ComboBox cboProfile;

		private QATool.CommObj          commObj = new CommObj();
        private AC_Counter              pstCountObj = new AC_Counter();

		private Outlook.Application		oApp;
		private Outlook._NameSpace		oNameSpace;
		private System.Data.DataSet		dsSubj;
		private System.Data.DataTable	dtSubject;
        private System.Windows.Forms.RadioButton rdoBinSearch;
        private System.Windows.Forms.RadioButton rdoSubStr;
		private Outlook.MAPIFolder		olPSTFolder;
        
		
		public ACTestPage()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

		}// end of ACTestPage

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			Debug.WriteLine("ACTestpage.cs - Dispose");
			if( disposing )
			{
				if(components != null)
				{
					Debug.WriteLine("\t Components Dispose");
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
            this.ttpACPage = new System.Windows.Forms.ToolTip(this.components);
            this.lnkGUIDFile = new System.Windows.Forms.LinkLabel();
            this.lnkPSTFile = new System.Windows.Forms.LinkLabel();
            this.lnkNSFFile = new System.Windows.Forms.LinkLabel();
            this.chkOutFile = new System.Windows.Forms.CheckBox();
            this.txtLogFileName = new System.Windows.Forms.TextBox();
            this.rdoBinSearch = new System.Windows.Forms.RadioButton();
            this.rdoSubStr = new System.Windows.Forms.RadioButton();
            this.btnPSTCheck = new System.Windows.Forms.Button();
            this.txtGuidFileName = new System.Windows.Forms.TextBox();
            this.txtPSTFileName = new System.Windows.Forms.TextBox();
            this.txtNSFFileName = new System.Windows.Forms.TextBox();
            this.btnNSFCheck = new System.Windows.Forms.Button();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.lsbDisplay = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lnkProfile = new System.Windows.Forms.LinkLabel();
            this.cboProfile = new System.Windows.Forms.ComboBox();
            this.dsSubj = new System.Data.DataSet();
            this.dtSubject = new System.Data.DataTable();
            ((System.ComponentModel.ISupportInitialize)(this.dsSubj)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtSubject)).BeginInit();
            this.SuspendLayout();
            // 
            // lnkGUIDFile
            // 
            this.lnkGUIDFile.Location = new System.Drawing.Point(8, 8);
            this.lnkGUIDFile.Name = "lnkGUIDFile";
            this.lnkGUIDFile.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lnkGUIDFile.Size = new System.Drawing.Size(60, 16);
            this.lnkGUIDFile.TabIndex = 3;
            this.lnkGUIDFile.TabStop = true;
            this.lnkGUIDFile.Text = "GUID File";
            this.lnkGUIDFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpACPage.SetToolTip(this.lnkGUIDFile, "Open GUID file");
            this.lnkGUIDFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkGUIDFile_LinkClicked);
            // 
            // lnkPSTFile
            // 
            this.lnkPSTFile.Location = new System.Drawing.Point(8, 32);
            this.lnkPSTFile.Name = "lnkPSTFile";
            this.lnkPSTFile.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lnkPSTFile.Size = new System.Drawing.Size(60, 16);
            this.lnkPSTFile.TabIndex = 2;
            this.lnkPSTFile.TabStop = true;
            this.lnkPSTFile.Text = "PST File";
            this.lnkPSTFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpACPage.SetToolTip(this.lnkPSTFile, "Open PST file");
            this.lnkPSTFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkPSTFile_LinkClicked);
            // 
            // lnkNSFFile
            // 
            this.lnkNSFFile.Location = new System.Drawing.Point(8, 180);
            this.lnkNSFFile.Name = "lnkNSFFile";
            this.lnkNSFFile.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lnkNSFFile.Size = new System.Drawing.Size(60, 16);
            this.lnkNSFFile.TabIndex = 6;
            this.lnkNSFFile.TabStop = true;
            this.lnkNSFFile.Text = "NSF File";
            this.lnkNSFFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpACPage.SetToolTip(this.lnkNSFFile, "Open PST file");
            this.lnkNSFFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkNSFFile_LinkClicked);
            // 
            // chkOutFile
            // 
            this.chkOutFile.Checked = true;
            this.chkOutFile.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOutFile.Location = new System.Drawing.Point(8, 84);
            this.chkOutFile.Name = "chkOutFile";
            this.chkOutFile.Size = new System.Drawing.Size(60, 16);
            this.chkOutFile.TabIndex = 9;
            this.chkOutFile.Text = "Log file";
            this.ttpACPage.SetToolTip(this.chkOutFile, "Write to log file");
            this.chkOutFile.CheckedChanged += new System.EventHandler(this.chkOutFile_CheckedChanged);
            // 
            // txtLogFileName
            // 
            this.txtLogFileName.Location = new System.Drawing.Point(72, 80);
            this.txtLogFileName.Name = "txtLogFileName";
            this.txtLogFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtLogFileName.Size = new System.Drawing.Size(234, 20);
            this.txtLogFileName.TabIndex = 10;
            this.txtLogFileName.Text = "AC_CheckLog.txt";
            this.ttpACPage.SetToolTip(this.txtLogFileName, "AC check result file name");
            // 
            // rdoBinSearch
            // 
            this.rdoBinSearch.Checked = true;
            this.rdoBinSearch.Location = new System.Drawing.Point(68, 104);
            this.rdoBinSearch.Name = "rdoBinSearch";
            this.rdoBinSearch.Size = new System.Drawing.Size(96, 16);
            this.rdoBinSearch.TabIndex = 74;
            this.rdoBinSearch.TabStop = true;
            this.rdoBinSearch.Text = "Binary Search";
            this.ttpACPage.SetToolTip(this.rdoBinSearch, "Subject MUST be GUID ONLY");
            // 
            // rdoSubStr
            // 
            this.rdoSubStr.Location = new System.Drawing.Point(184, 104);
            this.rdoSubStr.Name = "rdoSubStr";
            this.rdoSubStr.Size = new System.Drawing.Size(112, 16);
            this.rdoSubStr.TabIndex = 75;
            this.rdoSubStr.Text = "SubString Search";
            this.ttpACPage.SetToolTip(this.rdoSubStr, "Subject may more than GUID");
            // 
            // btnPSTCheck
            // 
            this.btnPSTCheck.Location = new System.Drawing.Point(308, 80);
            this.btnPSTCheck.Name = "btnPSTCheck";
            this.btnPSTCheck.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnPSTCheck.Size = new System.Drawing.Size(72, 20);
            this.btnPSTCheck.TabIndex = 5;
            this.btnPSTCheck.Text = "PST Check";
            this.btnPSTCheck.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPSTCheck.Click += new System.EventHandler(this.btnPSTCheck_Click);
            // 
            // txtGuidFileName
            // 
            this.txtGuidFileName.Location = new System.Drawing.Point(72, 8);
            this.txtGuidFileName.Name = "txtGuidFileName";
            this.txtGuidFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtGuidFileName.Size = new System.Drawing.Size(310, 20);
            this.txtGuidFileName.TabIndex = 4;
            this.txtGuidFileName.Text = "";
            // 
            // txtPSTFileName
            // 
            this.txtPSTFileName.Location = new System.Drawing.Point(70, 32);
            this.txtPSTFileName.Name = "txtPSTFileName";
            this.txtPSTFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtPSTFileName.Size = new System.Drawing.Size(312, 20);
            this.txtPSTFileName.TabIndex = 1;
            this.txtPSTFileName.Text = "";
            // 
            // txtNSFFileName
            // 
            this.txtNSFFileName.Location = new System.Drawing.Point(72, 180);
            this.txtNSFFileName.Name = "txtNSFFileName";
            this.txtNSFFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtNSFFileName.Size = new System.Drawing.Size(234, 20);
            this.txtNSFFileName.TabIndex = 7;
            this.txtNSFFileName.Text = "";
            // 
            // btnNSFCheck
            // 
            this.btnNSFCheck.Location = new System.Drawing.Point(312, 180);
            this.btnNSFCheck.Name = "btnNSFCheck";
            this.btnNSFCheck.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnNSFCheck.Size = new System.Drawing.Size(72, 20);
            this.btnNSFCheck.TabIndex = 8;
            this.btnNSFCheck.Text = "NSF Check";
            this.btnNSFCheck.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNSFCheck.Click += new System.EventHandler(this.btnNSFCheck_Click);
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(8, 212);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtOutput.Size = new System.Drawing.Size(372, 76);
            this.txtOutput.TabIndex = 11;
            this.txtOutput.Text = "";
            // 
            // lsbDisplay
            // 
            this.lsbDisplay.HorizontalScrollbar = true;
            this.lsbDisplay.Location = new System.Drawing.Point(4, 300);
            this.lsbDisplay.Name = "lsbDisplay";
            this.lsbDisplay.Size = new System.Drawing.Size(376, 121);
            this.lsbDisplay.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Yellow;
            this.label1.Location = new System.Drawing.Point(40, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(192, 23);
            this.label1.TabIndex = 13;
            this.label1.Text = "under construction";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(224, 56);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '+';
            this.txtPassword.Size = new System.Drawing.Size(156, 20);
            this.txtPassword.TabIndex = 72;
            this.txtPassword.Text = "password0";
            // 
            // lnkProfile
            // 
            this.lnkProfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkProfile.Location = new System.Drawing.Point(24, 60);
            this.lnkProfile.Name = "lnkProfile";
            this.lnkProfile.Size = new System.Drawing.Size(44, 16);
            this.lnkProfile.TabIndex = 71;
            this.lnkProfile.TabStop = true;
            this.lnkProfile.Text = "Profile :";
            this.lnkProfile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboProfile
            // 
            this.cboProfile.Items.AddRange(new object[] {
                                                            "duser11",
                                                            "duser51"});
            this.cboProfile.Location = new System.Drawing.Point(72, 56);
            this.cboProfile.Name = "cboProfile";
            this.cboProfile.Size = new System.Drawing.Size(148, 21);
            this.cboProfile.TabIndex = 73;
            this.cboProfile.Text = "zjournal";
            // 
            // dsSubj
            // 
            this.dsSubj.DataSetName = "dsMailSubj";
            this.dsSubj.Locale = new System.Globalization.CultureInfo("en-US");
            this.dsSubj.Tables.AddRange(new System.Data.DataTable[] {
                                                                        this.dtSubject});
            // 
            // dtSubject
            // 
            this.dtSubject.TableName = "tblSubject";
            // 
            // ACTestPage
            // 
            this.Controls.Add(this.rdoSubStr);
            this.Controls.Add(this.rdoBinSearch);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.lnkProfile);
            this.Controls.Add(this.cboProfile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lsbDisplay);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.txtLogFileName);
            this.Controls.Add(this.chkOutFile);
            this.Controls.Add(this.btnNSFCheck);
            this.Controls.Add(this.txtNSFFileName);
            this.Controls.Add(this.lnkNSFFile);
            this.Controls.Add(this.txtPSTFileName);
            this.Controls.Add(this.lnkGUIDFile);
            this.Controls.Add(this.txtGuidFileName);
            this.Controls.Add(this.lnkPSTFile);
            this.Controls.Add(this.btnPSTCheck);
            this.Name = "ACTestPage";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.dsSubj)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtSubject)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

		/// <summary>
		/// Create 2 stream reader object: 1) for guid file, 2) for PST file
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnPSTCheck_Click(object sender, System.EventArgs e)
		{
            commObj.LogToFile("ACTestPage.cs - btnPSTCheck_Click");

			Trace.WriteLine("ACTestpage.cs - btnPSTCheck_Click");
            btnPSTCheck.Enabled = false;
	        this.Cursor = Cursors.WaitCursor;

			if( txtPSTFileName.Text == "" )
			{
				MessageBox.Show( "Please select the PST file name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}

			if( txtGuidFileName.Text == "" )
			{
				MessageBox.Show( "Please select the GUID file name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}

            try
            {
                oApp = new Outlook.ApplicationClass();
                oNameSpace = oApp.GetNamespace("MAPI");

                oNameSpace.Logon( cboProfile.Text, txtPassword.Text, false, true );
                oNameSpace.AddStore( txtPSTFileName.Text );

                Outlook.MailItem mailItem;
 
                olPSTFolder = oNameSpace.PickFolder();
                if( olPSTFolder == null )
                {
                    MessageBox.Show( "Action Cancel - Please check all the input field.", "Warning" );
                    oNameSpace.RemoveStore( olPSTFolder );
                    oNameSpace.Logoff();
                    return;
                }

                int mailCount = olPSTFolder.Items.Count;
                pstCountObj.mailCount = olPSTFolder.Items.Count; // back up mail count

                ArrayList subjAL = new ArrayList();
                foreach( System.Object sysObj in olPSTFolder.Items )
                {

                    commObj.LogToFile("Within foreach - pst folder" + olPSTFolder.Items.ToString() );

                    mailItem = (Outlook.MailItem)sysObj;
                    subjAL.Add( mailItem.Subject );
                }//end of foreach

#if(DEBUG)
                for( int j = 0; j < mailCount; j++ )
                    Debug.WriteLine( "subject - " + subjAL[j].ToString() );
#endif

                if( rdoBinSearch.Checked )
                    SearchByBinary( txtGuidFileName.Text, ref subjAL );
                else
                    SearchBySubString( txtGuidFileName.Text, ref subjAL );

                string strSummary = "Mail in PST = " + pstCountObj.mailCount
                    + "\r\nGuid count = " + pstCountObj.guidCount
                    + "\r\nMail Found = " + pstCountObj.mailFound
                    + "\r\nMail Not Found = " + pstCountObj.mailNotFound;

                if( chkOutFile.Checked )
                    commObj.WriteLineByLine( txtLogFileName.Text, strSummary );

                txtOutput.Text = strSummary;
            }//end of try
            catch( IOException ioEx )
            {
                Debug.WriteLine( "ACTestPage.cs - btnPSTCheck_Click: " + ioEx.Message.ToString() );
                commObj.LogToFile( "ACTestPage.cs - btnPSTCheck_Click: " + ioEx.Message.ToString() );
            }//end of catch

            oNameSpace.RemoveStore( olPSTFolder );
            oNameSpace.Logoff();

            btnPSTCheck.Enabled = true;
            this.Cursor = Cursors.Default;
		}//end of btnPSTCheck_Click

        /// <summary>
        /// Use the Array List Binary search function to perform searching.
        /// 1) The list must be sorted.
        /// 2) Watch out the memory usage - will be huge.
        /// 3) The array list is pass by reference.
        /// </summary>
        /// <param name="fn"></param>
        /// <param name="myList"></param>
        public void SearchByBinary(string fn, ref ArrayList myList)
        {
            commObj.WriteLineByLine( txtLogFileName.Text, "Start SearchByBinary - " + DateTime.Now.ToUniversalTime() );
            myList.Sort(); // for binary search
            StreamReader srGuid = new StreamReader( fn );
           
            try
            {
                string strGuid = "";
                while( (strGuid = srGuid.ReadLine()) != null )
                {
                    Object myObj = strGuid;
                    int myIndex = myList.BinarySearch( myObj );
                    if( myIndex < 0 )
                    {
                        pstCountObj.mailNotFound++;
                        commObj.WriteLineByLine( txtLogFileName.Text, strGuid + "Not Found" );
                    }
                    else
                    {
                        pstCountObj.mailFound++;
                        commObj.WriteLineByLine( txtLogFileName.Text, strGuid );
                    }//end of else
                    pstCountObj.guidCount++;
                }//end of while       
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( "ACTestPage.cs - SearchByBinary: " + ex.Message.ToString() );
                commObj.LogToFile( "ACTestPage.cs - SearchByBinary: " + ex.Message.ToString() );
            }//end of catch

            srGuid.Close();
            commObj.WriteLineByLine( txtLogFileName.Text, "End SearchByBinary - " + DateTime.Now.ToUniversalTime() );
        }//end of SearchByBinary

        /// <summary>
        /// Sequencial search with removing found item.
        /// </summary>
        /// <param name="fn"></param>
        /// <param name="myList"></param>
        public void SearchBySubString( string fn, ref ArrayList myList )
        {
            commObj.WriteLineByLine( txtLogFileName.Text, "Start SearchBySubString - " + DateTime.Now.ToUniversalTime() );
            StreamReader srGuid = new StreamReader( fn );
            bool found = true;
            try
            {
                string strGuid = "";
                while( (strGuid = srGuid.ReadLine()) != null )
                {
                    foreach( string str in myList )
                    {
                        int idx = 0;
                        if( (idx = str.IndexOf(strGuid)) != -1 ) // -1 not found
                        {
                            found = true;
                            commObj.WriteLineByLine( txtLogFileName.Text, strGuid );
                            myList.Remove(str);
                            break;
                        }
                        else
                        {
                            found = false;
                            commObj.WriteLineByLine( txtLogFileName.Text, strGuid + "Not Found" );
                        }
                    }//end of foreach
                    if( found )
                        pstCountObj.mailFound++;
                    else
                        pstCountObj.mailNotFound++;

                    pstCountObj.guidCount++;
                }//end of while
            }
            catch( Exception ex )
            {
                Debug.WriteLine( "ACTestPage.cs - SearchByBinary: " + ex.Message.ToString() );
                commObj.LogToFile( "ACTestPage.cs - SearchByBinary: " + ex.Message.ToString() );
            }//end of catch

            srGuid.Close();
            commObj.WriteLineByLine( txtLogFileName.Text, "End SearchBySubString - " + DateTime.Now.ToUniversalTime() );
        }//end of SearchBySubString

		private void btnNSFCheck_Click(object sender, System.EventArgs e)
		{
			Trace.WriteLine("ACTestpage.cs - btnNSFCheck_Click");
			MessageBox.Show( "Under construction" );
		
		}//end of btnNSFCheck_Click

		private void lnkGUIDFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{		
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				if( ofDlg.FileName != "" )
				{
					txtGuidFileName.Text = ofDlg.FileName;
				}//end of if - open file name
			}// end of if - open file dialog										
		}//end of lnkGUIDFile_LinkClicked

		private void lnkPSTFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				if( ofDlg.FileName != "" )
				{
					txtPSTFileName.Text = ofDlg.FileName;
				}//end of if - open file name
			}// end of if - open file dialog												
		}//end of lnkPSTFile_LinkClicked

		private void lnkNSFFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;
			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				if( ofDlg.FileName != "" )
				{
					txtNSFFileName.Text = ofDlg.FileName;
				}//end of if - open file name
			}// end of if - open file dialog			
		}//end of lnkNSFFile_LinkClicked

		private void chkOutFile_CheckedChanged(object sender, System.EventArgs e)
		{
			if( chkOutFile.Checked )
			{
				txtLogFileName.Enabled = true;
				txtLogFileName.Text = "AC_CheckLog.txt";
			}
			else
			{
				txtLogFileName.Enabled = false;
				txtLogFileName.Text = "";
			}		
		}// end of chkOutFile_CheckedChanged

	}//end of class - ACTestPage

	public class AC_Counter
	{
		private int _mailCount = 0;
		private int _guidCount = 0;
		private int _mailFound = 0;		
		private int _mailNotFound = 0;

		public int mailCount
		{
			get
			{
				return _mailCount;
			}
			set
			{
				_mailCount = value;
			}
		}// end of property - mailCount

		public int guidCount
		{
			get
			{
				return _guidCount;
			}
			set
			{
				_guidCount = value;
			}
		}//end of guidCount

		public int mailFound
		{
			get
			{
				return _mailFound;
			}
			set
			{
				_mailFound = value;
			}
		}//end of mailFound

		public int mailNotFound
		{
			get
			{
				return _mailNotFound;
			}
			set
			{
				_mailNotFound = value;
			}
		}//end of mailNotFound
	}//end of class AC_Counter

}
