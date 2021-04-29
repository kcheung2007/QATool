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
	/// Summary description for PstCounter.
	/// </summary>
	public class PstCounter : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.Button btnDone;
        private System.Windows.Forms.TextBox txtTotalMail;
        private System.Windows.Forms.TextBox txtPstsSize;
        private System.Windows.Forms.ToolTip ttpPstWnd;
        private System.Windows.Forms.ComboBox cboProfile;
        private System.Windows.Forms.Button btnCountIt;
        private System.Windows.Forms.LinkLabel lnkPSTFile;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtPstFileName;
        private System.Windows.Forms.Label lblTotalSize;
        private System.Windows.Forms.Label lblTotalMail;
        private System.Windows.Forms.LinkLabel lnkProfile;
        private System.Windows.Forms.DataGrid dtgPstResult;

        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.Button btnSam;
        private System.Windows.Forms.Button btnAbort;

        private Outlook.Application oApp;
        private Outlook._NameSpace	oNameSpace;
        private int                 totalMail = 0;
        private long                totalPstSize = 0;
        private string []           fileNames;
        private Thread              countSpecialThread;                
        private QATool.CommObj      commObj = new CommObj();
        private System.Windows.Forms.ContextMenu dgContextMenu;

        // delegate - 
        private delegate void DelegateUpdateDataGrid( DataTable dataTable );
        private DelegateUpdateDataGrid m_delegateUpdateDataGrid;

		public PstCounter()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            // Initialize delegates
            m_delegateUpdateDataGrid = new DelegateUpdateDataGrid(this.UpdateDataGrid);
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
                if( countSpecialThread != null && countSpecialThread.IsAlive )
                {
                    this.KillPstInFolderThread();
                    commObj.LogToFile( "Thread.log", "   KillPstInFolderThread Killed" );
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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PstCounter));
            this.btnDone = new System.Windows.Forms.Button();
            this.txtTotalMail = new System.Windows.Forms.TextBox();
            this.txtPstsSize = new System.Windows.Forms.TextBox();
            this.ttpPstWnd = new System.Windows.Forms.ToolTip(this.components);
            this.cboProfile = new System.Windows.Forms.ComboBox();
            this.btnCountIt = new System.Windows.Forms.Button();
            this.lnkPSTFile = new System.Windows.Forms.LinkLabel();
            this.btnSam = new System.Windows.Forms.Button();
            this.btnAbort = new System.Windows.Forms.Button();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtPstFileName = new System.Windows.Forms.TextBox();
            this.lblTotalSize = new System.Windows.Forms.Label();
            this.lblTotalMail = new System.Windows.Forms.Label();
            this.lnkProfile = new System.Windows.Forms.LinkLabel();
            this.dtgPstResult = new System.Windows.Forms.DataGrid();
            this.dgContextMenu = new System.Windows.Forms.ContextMenu();
            ((System.ComponentModel.ISupportInitialize)(this.dtgPstResult)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDone
            // 
            this.btnDone.Enabled = false;
            this.btnDone.Location = new System.Drawing.Point(320, 52);
            this.btnDone.Name = "btnDone";
            this.btnDone.Size = new System.Drawing.Size(64, 20);
            this.btnDone.TabIndex = 93;
            this.btnDone.Text = "Clean me";
            this.ttpPstWnd.SetToolTip(this.btnDone, "Clean up the mess on your own risk");
            this.btnDone.Click += new System.EventHandler(this.btnDone_Click);
            // 
            // txtTotalMail
            // 
            this.txtTotalMail.Location = new System.Drawing.Point(60, 52);
            this.txtTotalMail.Name = "txtTotalMail";
            this.txtTotalMail.Size = new System.Drawing.Size(104, 20);
            this.txtTotalMail.TabIndex = 92;
            this.txtTotalMail.Text = "";
            // 
            // txtPstsSize
            // 
            this.txtPstsSize.Location = new System.Drawing.Point(228, 52);
            this.txtPstsSize.Name = "txtPstsSize";
            this.txtPstsSize.Size = new System.Drawing.Size(88, 20);
            this.txtPstsSize.TabIndex = 91;
            this.txtPstsSize.Text = "";
            // 
            // cboProfile
            // 
            this.cboProfile.Items.AddRange(new object[] {
                                                            "Lithium"});
            this.cboProfile.Location = new System.Drawing.Point(60, 4);
            this.cboProfile.Name = "cboProfile";
            this.cboProfile.Size = new System.Drawing.Size(108, 21);
            this.cboProfile.TabIndex = 87;
            this.cboProfile.Text = "pstProfile";
            this.ttpPstWnd.SetToolTip(this.cboProfile, "Outlook Profile name");
            // 
            // btnCountIt
            // 
            this.btnCountIt.Location = new System.Drawing.Point(320, 28);
            this.btnCountIt.Name = "btnCountIt";
            this.btnCountIt.Size = new System.Drawing.Size(64, 20);
            this.btnCountIt.TabIndex = 84;
            this.btnCountIt.Text = "Count";
            this.ttpPstWnd.SetToolTip(this.btnCountIt, "Count Mail inside PST");
            this.btnCountIt.Click += new System.EventHandler(this.btnCountIt_Click);
            // 
            // lnkPSTFile
            // 
            this.lnkPSTFile.Location = new System.Drawing.Point(4, 32);
            this.lnkPSTFile.Name = "lnkPSTFile";
            this.lnkPSTFile.Size = new System.Drawing.Size(56, 16);
            this.lnkPSTFile.TabIndex = 82;
            this.lnkPSTFile.TabStop = true;
            this.lnkPSTFile.Text = "PST File:";
            this.ttpPstWnd.SetToolTip(this.lnkPSTFile, "Browse PST Files");
            this.lnkPSTFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkPSTFile_LinkClicked);
            // 
            // btnSam
            // 
            this.btnSam.Location = new System.Drawing.Point(320, 4);
            this.btnSam.Name = "btnSam";
            this.btnSam.Size = new System.Drawing.Size(36, 20);
            this.btnSam.TabIndex = 94;
            this.btnSam.Text = "Sam";
            this.ttpPstWnd.SetToolTip(this.btnSam, "Cout PST more than 137...");
            this.btnSam.Click += new System.EventHandler(this.btnSam_Click);
            // 
            // btnAbort
            // 
            this.btnAbort.Enabled = false;
            this.btnAbort.Image = ((System.Drawing.Image)(resources.GetObject("btnAbort.Image")));
            this.btnAbort.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btnAbort.Location = new System.Drawing.Point(360, 0);
            this.btnAbort.Name = "btnAbort";
            this.btnAbort.Size = new System.Drawing.Size(24, 24);
            this.btnAbort.TabIndex = 95;
            this.ttpPstWnd.SetToolTip(this.btnAbort, "Abort the special count");
            this.btnAbort.Click += new System.EventHandler(this.btnAbort_Click);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(172, 4);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '+';
            this.txtPassword.Size = new System.Drawing.Size(144, 20);
            this.txtPassword.TabIndex = 88;
            this.txtPassword.Text = "password0";
            // 
            // txtPstFileName
            // 
            this.txtPstFileName.Location = new System.Drawing.Point(60, 28);
            this.txtPstFileName.Name = "txtPstFileName";
            this.txtPstFileName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtPstFileName.Size = new System.Drawing.Size(256, 20);
            this.txtPstFileName.TabIndex = 83;
            this.txtPstFileName.Text = "";
            // 
            // lblTotalSize
            // 
            this.lblTotalSize.Location = new System.Drawing.Point(164, 56);
            this.lblTotalSize.Name = "lblTotalSize";
            this.lblTotalSize.Size = new System.Drawing.Size(60, 12);
            this.lblTotalSize.TabIndex = 90;
            this.lblTotalSize.Text = "PSTs Size";
            this.lblTotalSize.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblTotalMail
            // 
            this.lblTotalMail.Location = new System.Drawing.Point(-4, 56);
            this.lblTotalMail.Name = "lblTotalMail";
            this.lblTotalMail.Size = new System.Drawing.Size(60, 12);
            this.lblTotalMail.TabIndex = 89;
            this.lblTotalMail.Text = "Total Mails";
            this.lblTotalMail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lnkProfile
            // 
            this.lnkProfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkProfile.Location = new System.Drawing.Point(8, 8);
            this.lnkProfile.Name = "lnkProfile";
            this.lnkProfile.Size = new System.Drawing.Size(44, 16);
            this.lnkProfile.TabIndex = 86;
            this.lnkProfile.TabStop = true;
            this.lnkProfile.Text = "Profile :";
            this.lnkProfile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dtgPstResult
            // 
            this.dtgPstResult.AlternatingBackColor = System.Drawing.Color.Gainsboro;
            this.dtgPstResult.BackColor = System.Drawing.Color.Silver;
            this.dtgPstResult.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.dtgPstResult.CaptionBackColor = System.Drawing.Color.DarkSlateBlue;
            this.dtgPstResult.CaptionFont = new System.Drawing.Font("Tahoma", 8F);
            this.dtgPstResult.CaptionForeColor = System.Drawing.Color.White;
            this.dtgPstResult.ContextMenu = this.dgContextMenu;
            this.dtgPstResult.DataMember = "";
            this.dtgPstResult.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dtgPstResult.FlatMode = true;
            this.dtgPstResult.ForeColor = System.Drawing.Color.Black;
            this.dtgPstResult.GridLineColor = System.Drawing.Color.White;
            this.dtgPstResult.HeaderBackColor = System.Drawing.Color.DarkGray;
            this.dtgPstResult.HeaderForeColor = System.Drawing.Color.Black;
            this.dtgPstResult.LinkColor = System.Drawing.Color.DarkSlateBlue;
            this.dtgPstResult.Location = new System.Drawing.Point(0, 76);
            this.dtgPstResult.Name = "dtgPstResult";
            this.dtgPstResult.ParentRowsBackColor = System.Drawing.Color.Black;
            this.dtgPstResult.ParentRowsForeColor = System.Drawing.Color.White;
            this.dtgPstResult.ReadOnly = true;
            this.dtgPstResult.SelectionBackColor = System.Drawing.Color.DarkSlateBlue;
            this.dtgPstResult.SelectionForeColor = System.Drawing.Color.White;
            this.dtgPstResult.Size = new System.Drawing.Size(388, 352);
            this.dtgPstResult.TabIndex = 85;
            // 
            // PstCounter
            // 
            this.Controls.Add(this.btnAbort);
            this.Controls.Add(this.btnSam);
            this.Controls.Add(this.txtTotalMail);
            this.Controls.Add(this.txtPstsSize);
            this.Controls.Add(this.cboProfile);
            this.Controls.Add(this.btnCountIt);
            this.Controls.Add(this.lnkPSTFile);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtPstFileName);
            this.Controls.Add(this.lblTotalSize);
            this.Controls.Add(this.lblTotalMail);
            this.Controls.Add(this.lnkProfile);
            this.Controls.Add(this.dtgPstResult);
            this.Controls.Add(this.btnDone);
            this.Name = "PstCounter";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.dtgPstResult)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

        private void lnkPSTFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            txtPstFileName.Text = "";

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.Multiselect = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                fileNames = ofDlg.FileNames;
                foreach( string str in fileNames )
                {
                    txtPstFileName.Text += ";" + str;
                }//end of foreach

                //check the first char
                string tmpStr = txtPstFileName.Text.ToString();
                if( tmpStr[0] == ';' )
                    txtPstFileName.Text = txtPstFileName.Text.Remove(0,1);
            }//end of if        
        }//end of lnkPSTFile_LinkClicked

        /// <summary>
        /// 1) Create Outlook App object.
        /// 2) Get initial folder count and clean them out.
        /// 3) Assume the default outlook folder name is "Personal Folders" which cannot be removed
        /// 4) Create PST table count.
        /// 5) Binding with the data grid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCountIt_Click(object sender, System.EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            btnCountIt.Enabled = false;
            btnDone.Enabled = false;
            btnSam.Enabled = false;

            // reset variable:
            totalMail = 0;
            totalPstSize = 0;
            txtTotalMail.Text = "";
            txtPstsSize.Text   = "";

            if( txtPstFileName.Text == "" )
            {
                txtPstFileName.Focus();
                txtPstFileName.BackColor = Color.DeepPink;

                btnCountIt.Enabled = true; // exiting - ie: re-enable the button and cursor.
                this.Cursor = Cursors.Default;
                return;
            }//end of if - exiting

            try
            {
                oApp = new Outlook.ApplicationClass();
                oNameSpace = oApp.GetNamespace("MAPI");
                oNameSpace.Logon( cboProfile.Text, txtPassword.Text, false, true );

                ////
                // remove all folder - fail if map drive disconnected
                // Assume the default folder name is @"\\Personal Folders", if not change it
                RemovePstStore(); // clean up before do anything

                // Create a DataTable and Bind it to the DataGrid
                DataTable A = CreatePstCountTable();  // btnCountIt
                dtgPstResult.DataSource = A;
                SizeColumns(dtgPstResult); // in btnCountIt

                txtTotalMail.Text = totalMail.ToString();
                double l = (double)totalPstSize/1048576.00;                
                txtPstsSize.Text = String.Format( "{0:F2}MB", l );
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message.ToString() + "\n"
                    + ex.GetType().ToString() + ex.StackTrace.ToString();
                commObj.LogToFile( msg );
                MessageBox.Show( "Check log: " + ex.Message.ToString(), "Error" );
            }//end of generic Exception
            finally
            {
                try
                {
                    oNameSpace.Logoff(); // both sucessful or fail
                    if( oApp != null )
                    {
                        Debug.WriteLine("\t Quit Outlook");
                        commObj.LogToFile( "Quit outlook" );
                        oApp.Quit();
                        Process [] localByName = Process.GetProcessesByName("outlook");
                        for( int n = 0; n < localByName.Length; n++ )
                            localByName[n].Kill();
                    }
                }//end of try - logoff outlook
                catch( Exception ex2 )
                {
                    string msg = ex2.Message.ToString() + "\n"
                        + ex2.GetType().ToString() + ex2.StackTrace.ToString();
                    commObj.LogToFile( msg );
                    MessageBox.Show( "Check log: " + ex2.Message.ToString(), "Outlook Logoff" );
                }//end of catch - mainly for outlook logoff and quit
            }//end of finally

            // enable the export table context menu
            dgContextMenu.MenuItems.Clear(); // clear all items in context menu
            dgContextMenu.MenuItems.Add( "Export Table", 
                new System.EventHandler(this.MnuExportTable) ); // create even handler

            btnCountIt.Enabled = true;
            btnDone.Enabled = true;
            btnSam.Enabled = true;
            this.Cursor = Cursors.Default;
        }// end of btnCountIt_Click

        /// <summary>
        /// Dynamic create a table to store pst file name and its message count
        /// Col 1: Pst File Name
        /// Col 2: Mail counts
        /// </summary>
        /// <returns></returns>
        private DataTable CreatePstCountTable()
        {
            DataTable	aTable = new DataTable("Pst Count Table");
            DataColumn	dtCol;
            DataRow		dtRow;

            // Create ID column and add to the DataTable.
            dtCol = new DataColumn();
            dtCol.DataType= System.Type.GetType("System.Int32");
            dtCol.ColumnName = "ID";
            dtCol.AutoIncrement = true;
            dtCol.Caption = "ID";
            dtCol.ReadOnly = true;
            dtCol.Unique = true;
            // Add the column to the DataColumnCollection.
            aTable.Columns.Add(dtCol);
 
            // Create PST File Name column and add to the DataTable.
            dtCol = new DataColumn();
            dtCol.DataType= System.Type.GetType("System.String");
            dtCol.ColumnName = "File Name";
            dtCol.AutoIncrement = false;
            dtCol.Caption = "PST File Name";
            dtCol.ReadOnly = true;
            dtCol.Unique = true;
            aTable.Columns.Add(dtCol);

            // Create Mail count column and add to the DataTable
            dtCol = new DataColumn();
            dtCol.DataType= System.Type.GetType("System.Int32");
            dtCol.ColumnName = "Mail Count";
            dtCol.AutoIncrement = false;
            dtCol.Caption = "Mail Count";
            dtCol.ReadOnly = true;
            dtCol.Unique = false;
            aTable.Columns.Add(dtCol);

            // Create Mail count column and add to the DataTable
            dtCol = new DataColumn();
            dtCol.DataType= System.Type.GetType("System.Int32");
            dtCol.ColumnName = "PST Size";
            dtCol.AutoIncrement = false;
            dtCol.Caption = "PST Size";
            dtCol.ReadOnly = true;
            dtCol.Unique = false;
            aTable.Columns.Add(dtCol);
 
            // Adding data
            int len = fileNames.Length;            
            int tmpMail = 0;
            long tmpSize = 0;

            int k = 0;
            for( int i = 0; i < len; i++ )
            {
                k++;
                if( k == 137 ) // 137 is limitation of outlook session 1000 1000
                {
                    RemovePstStore();
                    oNameSpace.Logoff(); // end the session                    
                    oApp.Quit();
                    Process [] localByName = Process.GetProcessesByName("outlook");
                    for( int n = 0; n < localByName.Length; n++ )
                        localByName[n].Kill();

                    System.Threading.Thread.Sleep( 3000 );
                    k = 0; //reset k
                    oApp = new Outlook.ApplicationClass();
                    oNameSpace = oApp.GetNamespace("MAPI");
                    oNameSpace.Logon( cboProfile.Text, txtPassword.Text, false, true );

                }//end of if - remove pst store

                dtRow = aTable.NewRow();
                dtRow["ID"] = i + 1; // start from 1
                dtRow["File Name"] = getShortFileName(fileNames[i]);

                tmpMail = getPstMailCount( fileNames[i] );
                totalMail += tmpMail;
                dtRow["Mail Count"] = tmpMail;

                tmpSize = getPstFileSize( fileNames[i] );
                totalPstSize += tmpSize;
                dtRow["PST Size"] = tmpSize;

                aTable.Rows.Add(dtRow);	

                // Update GUI - so user know it is NOT hang
                txtTotalMail.Text = totalMail.ToString();
                double l = (double)totalPstSize/1048576.00;                
                txtPstsSize.Text = String.Format( "{0:F2}MB", l );
                txtTotalMail.Refresh();
                txtPstsSize.Refresh();

                if( countSpecialThread != null && countSpecialThread.IsAlive )
                {
                    Debug.WriteLine("\t Update Datagrid" );
//                    IAsyncResult r = BeginInvoke(m_delegateUpdateDataGrid, new object[] {aTable} );
                }
                else
                {
                    dtgPstResult.DataSource = aTable;
                    dtgPstResult.Refresh();
                }                
            }//end of for
		
            return( aTable );		
        }// end of CreatePstCountTable

        /// <summary>
        /// This function will pass into delegate function.
        /// </summary>
        /// <param name="dataTable"></param>
        private void UpdateDataGrid( DataTable dataTable )
        {
            dtgPstResult.DataSource = dataTable;
            SizeColumns(dtgPstResult); // thd_process
            dtgPstResult.Refresh();
        }//end of UpdateDataGrid

        private int getPstMailCount( string fileName )
        {
            Debug.WriteLine("PstCheckerWnd.cd - getPstMailCount");
            Debug.WriteLine("\t file name = " + fileName );

            // save for reference - adding the pst into outlook
            // oNameSpace.AddStore( @"D:\\R1_00017.pst" );
            oNameSpace.AddStore( fileName );
                
            //            int count = oNameSpace.Folders.Count;
            Outlook.MAPIFolder olMapiFolder = oNameSpace.Folders.GetLast();
            Debug.WriteLine("\t MAPIFolder = " + olMapiFolder.Folders.ToString() );

            int mailCount = olMapiFolder.Items.Count;
            Debug.WriteLine("\t mail count = " + mailCount.ToString() );

            return mailCount;
        }//end of getPstMailCount

        private long getPstFileSize( string fileName )
        {
            FileInfo fileInfo = new FileInfo(fileName);
            return( fileInfo.Length );
        }

        /// <summary>
        /// Passing the full path file name, then extract the filename + ext.
        /// eg) passing D://abc/def/123.ext --> 123.ext
        /// </summary>
        /// <param name="fullFileName"></param>
        /// <returns>string: file name + ext</returns>
        public string getShortFileName( string fullFileName )
        {
            FileInfo fileInfo = new FileInfo(fullFileName);
            return( fileInfo.Name );
        }//end of getShortFileName

        private void RemovePstStore()
        {
            ////
            // remove all folder - fail if map drive disconnected
            // Assume the default folder name is @"\\Personal Folders", if not change it
            int numFolder = oNameSpace.Folders.Count; // folder of pst.. not inbox etc.            
            for( int i = 0; i < numFolder; i++ ) // start from 2nd one 
            {
                string path = oNameSpace.Folders.GetLast().FolderPath;
                Debug.WriteLine( "MAPI Folder = " + path );

                string str = @"\\Personal Folders"; // Assume the default folder name
                if( path == str )                   // if NOT, change it.
                    break;

                Outlook.MAPIFolder olMapiFolder = oNameSpace.Folders.GetLast();
                oNameSpace.RemoveStore( olMapiFolder );
            }//end of for

            numFolder = oNameSpace.Folders.Count; // folder of pst.. not inbox etc.            
            for( int i = 0; i < numFolder; i++ ) // start from 2nd one 
            {
                string path = oNameSpace.Folders.GetFirst().FolderPath;
                Debug.WriteLine( "MAPI Folder = " + path );

                string str = @"\\Personal Folders"; // Assume the default folder name
                if( path == str )                   // if NOT, change it.
                    break;

                Outlook.MAPIFolder olMapiFolder = oNameSpace.Folders.GetFirst();
                oNameSpace.RemoveStore( olMapiFolder );
            }//end of for
        }//end of RemovePstStore

        /// <summary>
        /// Clean up the added pst store. Map drive path must accessible
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDone_Click(object sender, System.EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            btnDone.Enabled = false;
            btnCountIt.Enabled = false;

            try
            {
                oApp = new Outlook.ApplicationClass();
                oNameSpace = oApp.GetNamespace("MAPI");
                oNameSpace.Logon( cboProfile.Text, txtPassword.Text, false, true );
                
                RemovePstStore();

                txtTotalMail.Text = "";
                txtPstsSize.Text = "";
                txtPstFileName.Text = "";

                DataTable clearDataTable = new DataTable();
                dtgPstResult.DataSource = clearDataTable;
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message.ToString() + "\n"
                    + ex.GetType().ToString() + ex.StackTrace.ToString();
                commObj.LogToFile( msg );
                MessageBox.Show( "Check log: " + ex.Message.ToString(), "Error" );
            }//end of generic Exception
            finally
            {
                try
                {
                    oNameSpace.Logoff(); // both sucessful or fail
                    if( oApp != null )
                    {
                        Debug.WriteLine("\t Quit Outlook");
                        commObj.LogToFile( "Quit outlook" );
                        oApp.Quit();
                        Process [] localByName = Process.GetProcessesByName("outlook");
                        for( int n = 0; n < localByName.Length; n++ )
                            localByName[n].Kill();
                    }
                }//end of try
                catch( Exception ex2 )
                {
                    string msg = ex2.Message.ToString() + "\n"
                        + ex2.GetType().ToString() + ex2.StackTrace.ToString();
                    commObj.LogToFile( msg );
                    MessageBox.Show( "Check log: " + ex2.Message.ToString(), "Outlook Logoff" );
                }//end of catch - mainly for outlook logoff and quit
            }//end of finally
            btnDone.Enabled = true;
            btnCountIt.Enabled = true;
            this.Cursor = Cursors.Default;
        }//end of btnDone_Click - Clean me

        /// <summary>
        /// Get the absolute path and export file name. Default file name is the table name .csv
        /// eg: table name is mytable, file name is mytable.csv
        /// </summary>
        /// <param name="fn">input default file name</param>
        /// <returns>string: full path filename</returns>
        public string GetSaveAbsPathFileName( string fn )
        {
            string filename = "PstCount.csv";
            SaveFileDialog saveFileDialog = new SaveFileDialog(); 
            saveFileDialog.Filter = "csv files (*.csv)|*.csv|txt files (*.txt)|*.txt|All files (*.*)|*.*"  ;
            saveFileDialog.FilterIndex = 1 ;
            saveFileDialog.RestoreDirectory = true ;
            saveFileDialog.FileName = fn;
 
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
                filename = saveFileDialog.FileName;

            return( filename );
        }// end of GetSaveAbsPathFileName

        /// <summary>
        /// Event handler for exporting whole table to CSV file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void MnuExportTable(System.Object sender, System.EventArgs e)
        {
            Debug.WriteLine("PstCounter.cs - MnuExportTable");

            // Create the CSV file to which grid data will be exported.
            string exportFileName = GetSaveAbsPathFileName( Application.StartupPath.ToString() );
            try
            {   DataTable m_dtDataTable = (DataTable)dtgPstResult.DataSource;

                StreamWriter sw = new StreamWriter( exportFileName );
                // First we will write the headers.
                int colCount = m_dtDataTable.Columns.Count;
                for( int i = 0; i < colCount; i++ )
                {
                    sw.Write( m_dtDataTable.Columns[i] );
                    if( i < colCount - 1 )
                        sw.Write(",");
                }//end of for
                sw.Write( sw.NewLine );

                // OK... Now write all the rows
                foreach( DataRow dataRow in m_dtDataTable.Rows )
                {
                    for(int i = 0; i < colCount; i++)
                    {
                        if( !Convert.IsDBNull( dataRow[i]) )
                        {
                            sw.Write(dataRow[i].ToString());
                        }
                        if( i < colCount - 1)
                            sw.Write(",");
                    }//end of for
                    sw.Write(sw.NewLine);
                }//end of foreach

                sw.Close();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch

        }//end of MnuExportTable

        /// <summary>
        /// Special case for Sam P. that counting the pst file more than 137.
        /// 137 is the limitation of outlook. Therefore need special handling.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSam_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine( "PstCount.cs - lnkFolder_LinkClicked" );

            countSpecialThread = new Thread( new ThreadStart(this.Thd_PstInFolder) );
            countSpecialThread.Name = "PstInFolderThread";
            countSpecialThread.Start();

            commObj.LogToFile( "Thread.log", "++ PstInFolderThread Start ++" ); 
      
        }//end of btnSam_Click

        private void Thd_PstInFolder()
        {
            Trace.WriteLine("PstCounter.cs - Thd_PstInFolder");

            string tmplable = lnkPSTFile.Text; // store it, restore later
            lnkPSTFile.Text = "Folder";
            btnAbort.Enabled = true;
            btnSam.Enabled   = false;

            FolderBrowserDialog fbDlg = new FolderBrowserDialog();

            fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
            if( txtPstFileName.Text != null )
                fbDlg.SelectedPath = Application.StartupPath.ToString();  // set the default folder

            if( fbDlg.ShowDialog() == DialogResult.OK )
            {
                txtPstFileName.Text = fbDlg.SelectedPath;
                txtPstFileName.Refresh();
            }
            else
            {
                lnkPSTFile.Text = tmplable; // restore it
                txtPstFileName.Text = "";
                btnAbort.Enabled = false;
                return; // do nothing
            }

            this.Cursor = Cursors.WaitCursor;
            btnCountIt.Enabled = false;
            btnDone.Enabled = false;
            // reset variable:
            totalMail = 0;
            totalPstSize = 0;
            txtTotalMail.Text = "";
            txtPstsSize.Text   = "";

            DirectoryInfo di = new DirectoryInfo(fbDlg.SelectedPath.ToString()); // attachment folder
            FileInfo[] lstFiles = di.GetFiles();
            fileNames = new string[lstFiles.Length];
            for( int i = 0; i < lstFiles.Length; i++ )
            {
                fileNames[i] = lstFiles[i].FullName;
            }//end of for
            
            try
            {
                oApp = new Outlook.ApplicationClass();
                oNameSpace = oApp.GetNamespace("MAPI");
                oNameSpace.Logon( cboProfile.Text, txtPassword.Text, false, true );

                RemovePstStore();

                // Create a DataTable and Bind it to the DataGrid
                DataTable A = CreatePstCountTable();  // thd_process
                // dtgPstResult.DataSource = A;
                IAsyncResult r = BeginInvoke(m_delegateUpdateDataGrid, new object[] {A} );

                txtTotalMail.Text = totalMail.ToString();
                double l = (double)totalPstSize/1048576.00;                
                txtPstsSize.Text = String.Format( "{0:F2}MB", l );
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message.ToString() + "\n"
                    + ex.GetType().ToString() + ex.StackTrace.ToString();
                commObj.LogToFile( msg );
                MessageBox.Show( "Check log: " + ex.Message.ToString(), "Error" );
            }//end of generic Exception
            finally
            {
                try
                {
                    oNameSpace.Logoff(); // both sucessful or fail
                    if( oApp != null )
                    {
                        commObj.LogToFile( "Quit outlook" );
                        oApp.Quit();
                    }
                }//end of try - logoff outlook
                catch( Exception ex2 )
                {
                    string msg = ex2.Message.ToString() + "\n"
                        + ex2.GetType().ToString() + ex2.StackTrace.ToString();
                    commObj.LogToFile( msg );
                    MessageBox.Show( "Check log: " + ex2.Message.ToString(), "Outlook Logoff" );
                }//end of catch - mainly for outlook logoff and quit
            }//end of finally

            btnCountIt.Enabled = true;
            btnDone.Enabled = true;
            this.Cursor = Cursors.Default;

            lnkPSTFile.Text = tmplable; // restore it
            txtPstFileName.Text = "";        
            btnAbort.Enabled = false;
            btnSam.Enabled   = true;

            // enable the export table context menu
            dgContextMenu.MenuItems.Clear(); // clear all items in context menu
            dgContextMenu.MenuItems.Add( "Export Table", 
                new System.EventHandler(this.MnuExportTable) ); // create even handler

        }//end of Thd_PstInFolder

        private void KillPstInFolderThread()
        {
            Trace.WriteLine("PstCounter.cs - KillPstInFolderThread()");
            try
            {
                commObj.LogToFile( "Thread.log", "Kill KillPstInFolderThread Start");
                countSpecialThread.Abort(); // abort
                countSpecialThread.Join();  // require for ensure the thread kill
            }//end of try 
            catch( ThreadAbortException thdEx )
            {
                Trace.WriteLine( thdEx.Message );
                commObj.LogToFile( "Aborting the PstInFolder thread : " + thdEx.Message.ToString() );
            }//end of catch				

        }//end of KillPstInFolderThread

        /// <summary>
        /// Auto Size the column of datagrid
        /// </summary>
        /// <param name="grid"></param>
        protected void SizeColumns(DataGrid grid)
        {
            Debug.WriteLine("PstCounter.cs - SizeColumns()" );
//            Graphics g = CreateGraphics();  
            try
            {
                DataTable dataTable = (DataTable)grid.DataSource;

                bool isContained = grid.TableStyles.Contains( dataTable.TableName );
                if( isContained )
                    return;

                DataGridTableStyle dataGridTableStyle = new DataGridTableStyle();
                dataGridTableStyle.MappingName = dataTable.TableName;
                dataGridTableStyle.AlternatingBackColor = Color.Gainsboro;

                Debug.WriteLine("\t Table Name = " + dataTable.TableName );

                Graphics g = CreateGraphics();
                foreach(DataColumn dataColumn in dataTable.Columns)
                {
                    int maxSize = 0;

                    SizeF size = g.MeasureString( dataColumn.ColumnName, grid.Font );
                    if( maxSize < size.Width)
                        maxSize = (int)size.Width;

                    foreach(DataRow row in dataTable.Rows)
                    {
                        size = g.MeasureString( row[dataColumn.ColumnName].ToString(), grid.Font );

                        if( maxSize < size.Width )
                            maxSize = (int)size.Width;
                    }// end of foreach

                    DataGridColumnStyle dataGridColumnStyle =  new DataGridTextBoxColumn();
                    dataGridColumnStyle.MappingName = dataColumn.ColumnName;
                    dataGridColumnStyle.HeaderText = dataColumn.ColumnName;
                    dataGridColumnStyle.Width = maxSize + 5;
                    dataGridTableStyle.GridColumnStyles.Add(dataGridColumnStyle);
                }// end of foreach
                g.Dispose();
                grid.TableStyles.Add(dataGridTableStyle);
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
            finally
            {
                Debug.WriteLine("\t finally end of SizeColumns" );
//                g.Dispose();
            }//end of finally                      
        }//end of SizeColumns

        private void btnAbort_Click(object sender, System.EventArgs e)
        {
            try
            {
                btnAbort.Enabled = false;
                btnCountIt.Enabled = true;
                btnSam.Enabled = true;
                this.Cursor = Cursors.Default;

                if( countSpecialThread != null && countSpecialThread.IsAlive )
                    this.KillPstInFolderThread();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine("PstCounter.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                commObj.LogToFile("PstCounter.cs - btnAbort_Click " + ex.Message + "\n" + ex.StackTrace );
                MessageBox.Show( ex.Message + "\n" + ex.StackTrace, "Abort Exception" );
            }//end of catch        
        }

	}
}
