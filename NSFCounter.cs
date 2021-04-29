using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for NSFCounter.
	/// </summary>
	public class NSFCounter : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.ToolTip ttpNSFCounter;
        private System.Windows.Forms.DataGrid dtgNsfResult;
        private System.Windows.Forms.LinkLabel lnkNSFFiles;
        private System.Windows.Forms.Label lblIDPassword;
        private System.Windows.Forms.TextBox txtLoginID;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtTotalMail;
        private System.Windows.Forms.TextBox txtNsfsSize;
        private System.Windows.Forms.Label lblTotalSize;
        private System.Windows.Forms.Label lblTotalMail;
        private System.Windows.Forms.Button btnCount;
        private System.ComponentModel.IContainer components;

        private int      totalMail = 0;
        private double   totalNsfSize = 0;
        private System.Windows.Forms.Label lblStatus;
        private string[] fileNames;

		public NSFCounter()
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
            this.ttpNSFCounter = new System.Windows.Forms.ToolTip(this.components);
            this.lnkNSFFiles = new System.Windows.Forms.LinkLabel();
            this.txtLoginID = new System.Windows.Forms.TextBox();
            this.btnCount = new System.Windows.Forms.Button();
            this.dtgNsfResult = new System.Windows.Forms.DataGrid();
            this.lblIDPassword = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtTotalMail = new System.Windows.Forms.TextBox();
            this.txtNsfsSize = new System.Windows.Forms.TextBox();
            this.lblTotalSize = new System.Windows.Forms.Label();
            this.lblTotalMail = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgNsfResult)).BeginInit();
            this.SuspendLayout();
            // 
            // lnkNSFFiles
            // 
            this.lnkNSFFiles.Location = new System.Drawing.Point(4, 56);
            this.lnkNSFFiles.Name = "lnkNSFFiles";
            this.lnkNSFFiles.Size = new System.Drawing.Size(60, 16);
            this.lnkNSFFiles.TabIndex = 1;
            this.lnkNSFFiles.TabStop = true;
            this.lnkNSFFiles.Text = "NSF Files:";
            this.lnkNSFFiles.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ttpNSFCounter.SetToolTip(this.lnkNSFFiles, "Browse NSF Files");
            this.lnkNSFFiles.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkNSFFiles_LinkClicked);
            // 
            // txtLoginID
            // 
            this.txtLoginID.Enabled = false;
            this.txtLoginID.Location = new System.Drawing.Point(72, 4);
            this.txtLoginID.Name = "txtLoginID";
            this.txtLoginID.Size = new System.Drawing.Size(144, 20);
            this.txtLoginID.TabIndex = 3;
            this.txtLoginID.Text = "";
            this.ttpNSFCounter.SetToolTip(this.txtLoginID, "Notes Login ID");
            // 
            // btnCount
            // 
            this.btnCount.Enabled = false;
            this.btnCount.Location = new System.Drawing.Point(304, 52);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 20);
            this.btnCount.TabIndex = 97;
            this.btnCount.Text = "Count";
            this.ttpNSFCounter.SetToolTip(this.btnCount, "Count mail inside NSF");
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // dtgNsfResult
            // 
            this.dtgNsfResult.AlternatingBackColor = System.Drawing.Color.LightGray;
            this.dtgNsfResult.BackColor = System.Drawing.Color.DarkGray;
            this.dtgNsfResult.CaptionBackColor = System.Drawing.Color.White;
            this.dtgNsfResult.CaptionFont = new System.Drawing.Font("Verdana", 10F);
            this.dtgNsfResult.CaptionForeColor = System.Drawing.Color.Navy;
            this.dtgNsfResult.DataMember = "";
            this.dtgNsfResult.ForeColor = System.Drawing.Color.Black;
            this.dtgNsfResult.GridLineColor = System.Drawing.Color.Black;
            this.dtgNsfResult.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None;
            this.dtgNsfResult.HeaderBackColor = System.Drawing.Color.Silver;
            this.dtgNsfResult.HeaderForeColor = System.Drawing.Color.Black;
            this.dtgNsfResult.LinkColor = System.Drawing.Color.Navy;
            this.dtgNsfResult.Location = new System.Drawing.Point(8, 76);
            this.dtgNsfResult.Name = "dtgNsfResult";
            this.dtgNsfResult.ParentRowsBackColor = System.Drawing.Color.White;
            this.dtgNsfResult.ParentRowsForeColor = System.Drawing.Color.Black;
            this.dtgNsfResult.SelectionBackColor = System.Drawing.Color.Navy;
            this.dtgNsfResult.SelectionForeColor = System.Drawing.Color.White;
            this.dtgNsfResult.Size = new System.Drawing.Size(376, 348);
            this.dtgNsfResult.TabIndex = 0;
            // 
            // lblIDPassword
            // 
            this.lblIDPassword.Enabled = false;
            this.lblIDPassword.Location = new System.Drawing.Point(4, 8);
            this.lblIDPassword.Name = "lblIDPassword";
            this.lblIDPassword.Size = new System.Drawing.Size(68, 12);
            this.lblIDPassword.TabIndex = 2;
            this.lblIDPassword.Text = "ID/Password";
            this.lblIDPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtPassword
            // 
            this.txtPassword.Enabled = false;
            this.txtPassword.Location = new System.Drawing.Point(220, 4);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '+';
            this.txtPassword.Size = new System.Drawing.Size(160, 20);
            this.txtPassword.TabIndex = 4;
            this.txtPassword.Text = "password0";
            // 
            // txtTotalMail
            // 
            this.txtTotalMail.Location = new System.Drawing.Point(72, 28);
            this.txtTotalMail.Name = "txtTotalMail";
            this.txtTotalMail.Size = new System.Drawing.Size(120, 20);
            this.txtTotalMail.TabIndex = 96;
            this.txtTotalMail.Text = "";
            // 
            // txtNsfsSize
            // 
            this.txtNsfsSize.Location = new System.Drawing.Point(256, 28);
            this.txtNsfsSize.Name = "txtNsfsSize";
            this.txtNsfsSize.Size = new System.Drawing.Size(124, 20);
            this.txtNsfsSize.TabIndex = 95;
            this.txtNsfsSize.Text = "";
            // 
            // lblTotalSize
            // 
            this.lblTotalSize.Location = new System.Drawing.Point(192, 32);
            this.lblTotalSize.Name = "lblTotalSize";
            this.lblTotalSize.Size = new System.Drawing.Size(60, 12);
            this.lblTotalSize.TabIndex = 94;
            this.lblTotalSize.Text = "NSFs Size";
            this.lblTotalSize.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblTotalMail
            // 
            this.lblTotalMail.Location = new System.Drawing.Point(8, 32);
            this.lblTotalMail.Name = "lblTotalMail";
            this.lblTotalMail.Size = new System.Drawing.Size(60, 12);
            this.lblTotalMail.TabIndex = 93;
            this.lblTotalMail.Text = "Total Mails";
            this.lblTotalMail.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblStatus
            // 
            this.lblStatus.Location = new System.Drawing.Point(72, 56);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(228, 16);
            this.lblStatus.TabIndex = 98;
            // 
            // NSFCounter
            // 
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.txtTotalMail);
            this.Controls.Add(this.txtNsfsSize);
            this.Controls.Add(this.lblTotalSize);
            this.Controls.Add(this.lblTotalMail);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtLoginID);
            this.Controls.Add(this.lblIDPassword);
            this.Controls.Add(this.lnkNSFFiles);
            this.Controls.Add(this.dtgNsfResult);
            this.Name = "NSFCounter";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.dtgNsfResult)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

        private void lnkNSFFiles_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.Multiselect = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                fileNames = ofDlg.FileNames;
                btnCount.Enabled = true;

                // Create a DataTable and Bind it to the DataGrid
                DataTable nsfTable = CreateNsfCountTable();

                // Adding data - file names
                int len = fileNames.Length;            
                DataRow		dtRow;            
                for( int i = 0; i < len; i++ )
                {
                    dtRow = nsfTable.NewRow();
                    dtRow["ID"] = i + 1; // start from 1
                    dtRow["File Name"] = getShortFileName(fileNames[i]);

                    nsfTable.Rows.Add(dtRow);	
                }//end of for

                dtgNsfResult.DataSource = nsfTable;
            }//end of if - get the nsf names
        }//end of lnkNSFFiles_LinkClicked

        private DataTable CreateNsfCountTable()
        {
            DataTable	aTable = new DataTable("NSF Count Table");
            DataColumn	dtCol;

            try
            {
                // Create ID column and add to the DataTable.
                dtCol = new DataColumn();
                dtCol.DataType= System.Type.GetType("System.Int32");
                dtCol.ColumnName = "ID";
                dtCol.AutoIncrement = true;
                dtCol.Caption = "ID";
                dtCol.Unique = true;
                // Add the column to the DataColumnCollection.
                aTable.Columns.Add(dtCol);
 
                // Create PST File Name column and add to the DataTable.
                dtCol = new DataColumn();
                dtCol.DataType= System.Type.GetType("System.String");
                dtCol.ColumnName = "File Name";
                dtCol.AutoIncrement = false;
                dtCol.Caption = "NSF File Name";
                //dtCol.ReadOnly = true;
                dtCol.Unique = true;
                aTable.Columns.Add(dtCol);

                // Create Mail count column and add to the DataTable
                dtCol = new DataColumn();
                dtCol.DataType= System.Type.GetType("System.Int32");
                dtCol.ColumnName = "Mail Count";
                dtCol.AutoIncrement = false;
                dtCol.Caption = "Mail Count";
                dtCol.Unique = false;
                aTable.Columns.Add(dtCol);

                // Create Mail count column and add to the DataTable
                dtCol = new DataColumn();
                dtCol.DataType= System.Type.GetType("System.Int32");
                dtCol.ColumnName = "NSF Size";
                dtCol.AutoIncrement = false;
                dtCol.Caption = "NSF Size";
                dtCol.Unique = false;
                aTable.Columns.Add(dtCol);
                
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace;
                MessageBox.Show( msg );
            }//end of catch - exception

            return( aTable );
        }// end of CreateNsfCountTable

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

        private void btnCount_Click(object sender, System.EventArgs e)
        {
            CleanUpUICtrl(); 

            // Adding data - size and mail count
            int     len = fileNames.Length;            
            int     tmpMail = 0;
            double  tmpSize = 0;

            try
            {
                Domino.NotesSession domSession = new Domino.NotesSession();

                domSession.Initialize("");
                // Used by install client only - Save next line for reference
                //domSession.InitializeUsingNotesUserName("admin", "password0");

                txtLoginID.Text = domSession.UserName;
//                txtPassword.Text = domSession.ServerName;

                // after initialize session, get the db open
                Domino.NotesDatabase domDB;
                DataTable myTable = (DataTable)dtgNsfResult.DataSource;
                for( int i = 0; i < len; i++ )
                {
                    domDB = domSession.GetDatabase("", fileNames[i], false);
                    if( domDB == null )
                        continue; // move to next file

                    if( domDB.IsOpen )
                    {
                        lblStatus.Text = domDB.Title;
                        tmpMail = domDB.AllDocuments.Count;
                        myTable.Rows[i]["Mail Count"] = tmpMail;
                        totalMail += tmpMail;
                        tmpSize = domDB.Size;
                        myTable.Rows[i]["NSF Size"] = tmpSize;
                        totalNsfSize += tmpSize;

                    }//end of if
                    else
                    {
                        MessageBox.Show( "Error - Fail to Open nsf file " + fileNames[i].ToString() );
                    }//end of else

                    dtgNsfResult.ReadOnly = true;
                    dtgNsfResult.Update();

                    txtTotalMail.Text = totalMail.ToString();
                    double l = (double)totalNsfSize/1048576.00;
                    txtNsfsSize.Text  = String.Format( "{0:F2}MB", l );

                }//end of foreach
            }//end of try
            catch( Exception ex )
            {
                string msg = ex.Message + "\n" + ex.GetType().ToString() + "\n" + ex.StackTrace;
                MessageBox.Show( msg );
            }//end of catch       
        }//end of btnCount_Click

        /// <summary>
        /// Basically, reset the default control value - null
        /// </summary>
        public void CleanUpUICtrl()
        {
            txtLoginID.Text   = "";
            txtTotalMail.Text = "";
            txtNsfsSize.Text  = "";
            lblStatus.Text    = "";

            totalMail    = 0;
            totalNsfSize = 0;        
        }
	}
}
