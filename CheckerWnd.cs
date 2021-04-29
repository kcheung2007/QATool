using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for CheckerWnd.
	/// Several Methods modified from http://authors.aspalliance.com
	/// </summary>
	public class CheckerWnd : System.Windows.Forms.Form
	{
        const int RESULT_WND = 1;
        const int SOURCE_WND = 0;

        private System.Windows.Forms.LinkLabel lnkSourceFile;
        private System.Windows.Forms.ToolTip ttpChecker;
        private System.Windows.Forms.LinkLabel lnkResultFile;
        private System.Windows.Forms.DataGrid dgSource;
        private System.Windows.Forms.DataGrid dgResult;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.RichTextBox rtbSource;
        private System.Windows.Forms.RichTextBox rtbResult;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.Button btnDeDupSrc;
        private System.Windows.Forms.Button btnGetSrcDup;
        private System.Windows.Forms.Button btnGetRsuDup;
        private System.Windows.Forms.Button btnDeDupRsu;
        private System.Windows.Forms.Button btnRefreshSrc;
        private System.Windows.Forms.Button btnRefreshRsu;

        private QATool.CommObj commObj = new CommObj();

        private int m_numMailSource = 0;
        private int m_numMailResult = 0;
        private int m_sourceDiffCnt = 0;
        private int m_resultDiffCnt = 0;
        private int m_sourceDeDupCnt = 0;
        private int m_resultDeDupCnt = 0;

        private string m_strMsg = "";

        private string[] m_strArrSource = null;
        private string[] m_strArrResult = null;

		private System.Windows.Forms.Button btnSrcCompare;
		private System.Windows.Forms.Button btnRsuCompare;
		private System.Windows.Forms.ToolBar toolBar1;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBar toolBar2;
        private System.Windows.Forms.Button btnReload2;
        private System.Windows.Forms.Button btnReload1;
        private System.Windows.Forms.ToolBarButton tbrBtnSave1;
        private System.Windows.Forms.ToolBarButton tbrBtnSave2;
        private System.Windows.Forms.ToolBarButton tbrBtnExport2;
        private System.Windows.Forms.ToolBarButton tbrBtnExport1;
        
        
        

		public CheckerWnd()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CheckerWnd));
            this.lnkSourceFile = new System.Windows.Forms.LinkLabel();
            this.ttpChecker = new System.Windows.Forms.ToolTip(this.components);
            this.lnkResultFile = new System.Windows.Forms.LinkLabel();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.rtbSource = new System.Windows.Forms.RichTextBox();
            this.rtbResult = new System.Windows.Forms.RichTextBox();
            this.btnDeDupSrc = new System.Windows.Forms.Button();
            this.btnGetSrcDup = new System.Windows.Forms.Button();
            this.btnGetRsuDup = new System.Windows.Forms.Button();
            this.btnDeDupRsu = new System.Windows.Forms.Button();
            this.btnRefreshSrc = new System.Windows.Forms.Button();
            this.btnRefreshRsu = new System.Windows.Forms.Button();
            this.btnSrcCompare = new System.Windows.Forms.Button();
            this.btnRsuCompare = new System.Windows.Forms.Button();
            this.btnReload2 = new System.Windows.Forms.Button();
            this.btnReload1 = new System.Windows.Forms.Button();
            this.dgSource = new System.Windows.Forms.DataGrid();
            this.dgResult = new System.Windows.Forms.DataGrid();
            this.toolBar1 = new System.Windows.Forms.ToolBar();
            this.tbrBtnSave1 = new System.Windows.Forms.ToolBarButton();
            this.tbrBtnExport1 = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.toolBar2 = new System.Windows.Forms.ToolBar();
            this.tbrBtnSave2 = new System.Windows.Forms.ToolBarButton();
            this.tbrBtnExport2 = new System.Windows.Forms.ToolBarButton();
            ((System.ComponentModel.ISupportInitialize)(this.dgSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).BeginInit();
            this.SuspendLayout();
            // 
            // lnkSourceFile
            // 
            this.lnkSourceFile.Location = new System.Drawing.Point(4, 8);
            this.lnkSourceFile.Name = "lnkSourceFile";
            this.lnkSourceFile.Size = new System.Drawing.Size(76, 16);
            this.lnkSourceFile.TabIndex = 0;
            this.lnkSourceFile.TabStop = true;
            this.lnkSourceFile.Text = "Source ID File";
            this.ttpChecker.SetToolTip(this.lnkSourceFile, "Browse for the source file contains GUID");
            this.lnkSourceFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkSourceFile_LinkClicked);
            // 
            // lnkResultFile
            // 
            this.lnkResultFile.Location = new System.Drawing.Point(296, 8);
            this.lnkResultFile.Name = "lnkResultFile";
            this.lnkResultFile.Size = new System.Drawing.Size(76, 16);
            this.lnkResultFile.TabIndex = 1;
            this.lnkResultFile.TabStop = true;
            this.lnkResultFile.Text = "Result ID File";
            this.ttpChecker.SetToolTip(this.lnkResultFile, "Browse for the result file contains GUID");
            this.lnkResultFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkResultFile_LinkClicked);
            // 
            // txtSource
            // 
            this.txtSource.Location = new System.Drawing.Point(84, 4);
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(200, 20);
            this.txtSource.TabIndex = 4;
            this.txtSource.Text = "";
            this.ttpChecker.SetToolTip(this.txtSource, "Full path of source file");
            this.txtSource.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.KeyPress_txtSource);
            // 
            // txtResult
            // 
            this.txtResult.Location = new System.Drawing.Point(372, 4);
            this.txtResult.Name = "txtResult";
            this.txtResult.Size = new System.Drawing.Size(200, 20);
            this.txtResult.TabIndex = 5;
            this.txtResult.Text = "";
            this.ttpChecker.SetToolTip(this.txtResult, "Full path of result file");
            this.txtResult.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.KeyPress_txtResult);
            // 
            // rtbSource
            // 
            this.rtbSource.Location = new System.Drawing.Point(84, 28);
            this.rtbSource.Name = "rtbSource";
            this.rtbSource.ReadOnly = true;
            this.rtbSource.Size = new System.Drawing.Size(200, 108);
            this.rtbSource.TabIndex = 6;
            this.rtbSource.Text = "rtbSource";
            this.ttpChecker.SetToolTip(this.rtbSource, "Display window");
            // 
            // rtbResult
            // 
            this.rtbResult.Location = new System.Drawing.Point(372, 28);
            this.rtbResult.Name = "rtbResult";
            this.rtbResult.ReadOnly = true;
            this.rtbResult.Size = new System.Drawing.Size(200, 108);
            this.rtbResult.TabIndex = 7;
            this.rtbResult.Text = "rtbResult";
            this.ttpChecker.SetToolTip(this.rtbResult, "Display window");
            // 
            // btnDeDupSrc
            // 
            this.btnDeDupSrc.Enabled = false;
            this.btnDeDupSrc.Location = new System.Drawing.Point(4, 52);
            this.btnDeDupSrc.Name = "btnDeDupSrc";
            this.btnDeDupSrc.Size = new System.Drawing.Size(75, 20);
            this.btnDeDupSrc.TabIndex = 9;
            this.btnDeDupSrc.Text = "De-Dups";
            this.ttpChecker.SetToolTip(this.btnDeDupSrc, "Remove Duplicated Items. Return an unique list.");
            this.btnDeDupSrc.Click += new System.EventHandler(this.btnDeDupSrc_Click);
            // 
            // btnGetSrcDup
            // 
            this.btnGetSrcDup.Enabled = false;
            this.btnGetSrcDup.Location = new System.Drawing.Point(4, 73);
            this.btnGetSrcDup.Name = "btnGetSrcDup";
            this.btnGetSrcDup.Size = new System.Drawing.Size(75, 20);
            this.btnGetSrcDup.TabIndex = 10;
            this.btnGetSrcDup.Text = "Get Dups";
            this.ttpChecker.SetToolTip(this.btnGetSrcDup, "Extract Duplicated Items");
            this.btnGetSrcDup.Click += new System.EventHandler(this.btnGetSrcDup_Click);
            // 
            // btnGetRsuDup
            // 
            this.btnGetRsuDup.Enabled = false;
            this.btnGetRsuDup.Location = new System.Drawing.Point(288, 73);
            this.btnGetRsuDup.Name = "btnGetRsuDup";
            this.btnGetRsuDup.Size = new System.Drawing.Size(75, 20);
            this.btnGetRsuDup.TabIndex = 11;
            this.btnGetRsuDup.Text = "Get Dups";
            this.ttpChecker.SetToolTip(this.btnGetRsuDup, "Extract Duplicated Items");
            this.btnGetRsuDup.Click += new System.EventHandler(this.btnGetRsuDup_Click);
            // 
            // btnDeDupRsu
            // 
            this.btnDeDupRsu.Enabled = false;
            this.btnDeDupRsu.Location = new System.Drawing.Point(288, 52);
            this.btnDeDupRsu.Name = "btnDeDupRsu";
            this.btnDeDupRsu.Size = new System.Drawing.Size(75, 20);
            this.btnDeDupRsu.TabIndex = 12;
            this.btnDeDupRsu.Text = "De-Dups";
            this.ttpChecker.SetToolTip(this.btnDeDupRsu, "Remove Duplicated Items. Return an unique list.");
            this.btnDeDupRsu.Click += new System.EventHandler(this.btnDeDupRsu_Click);
            // 
            // btnRefreshSrc
            // 
            this.btnRefreshSrc.Enabled = false;
            this.btnRefreshSrc.Location = new System.Drawing.Point(4, 115);
            this.btnRefreshSrc.Name = "btnRefreshSrc";
            this.btnRefreshSrc.Size = new System.Drawing.Size(75, 20);
            this.btnRefreshSrc.TabIndex = 13;
            this.btnRefreshSrc.Text = "Refresh";
            this.ttpChecker.SetToolTip(this.btnRefreshSrc, "Refresh data from memory");
            this.btnRefreshSrc.Click += new System.EventHandler(this.btnRefreshSrc_Click);
            // 
            // btnRefreshRsu
            // 
            this.btnRefreshRsu.Enabled = false;
            this.btnRefreshRsu.Location = new System.Drawing.Point(288, 115);
            this.btnRefreshRsu.Name = "btnRefreshRsu";
            this.btnRefreshRsu.Size = new System.Drawing.Size(75, 20);
            this.btnRefreshRsu.TabIndex = 14;
            this.btnRefreshRsu.Text = "Refresh";
            this.ttpChecker.SetToolTip(this.btnRefreshRsu, "Refresh data from memory");
            this.btnRefreshRsu.Click += new System.EventHandler(this.btnRefreshRsu_Click);
            // 
            // btnSrcCompare
            // 
            this.btnSrcCompare.Enabled = false;
            this.btnSrcCompare.Location = new System.Drawing.Point(4, 94);
            this.btnSrcCompare.Name = "btnSrcCompare";
            this.btnSrcCompare.Size = new System.Drawing.Size(75, 20);
            this.btnSrcCompare.TabIndex = 15;
            this.btnSrcCompare.Text = "Compare";
            this.ttpChecker.SetToolTip(this.btnSrcCompare, "Compare the source array against result aray");
            this.btnSrcCompare.Click += new System.EventHandler(this.btnSrcCompare_Click);
            // 
            // btnRsuCompare
            // 
            this.btnRsuCompare.Enabled = false;
            this.btnRsuCompare.Location = new System.Drawing.Point(288, 94);
            this.btnRsuCompare.Name = "btnRsuCompare";
            this.btnRsuCompare.Size = new System.Drawing.Size(75, 20);
            this.btnRsuCompare.TabIndex = 16;
            this.btnRsuCompare.Text = "Compare";
            this.ttpChecker.SetToolTip(this.btnRsuCompare, "Compare the result array against source aray");
            this.btnRsuCompare.Click += new System.EventHandler(this.btnRsuCompare_Click);
            // 
            // btnReload2
            // 
            this.btnReload2.Enabled = false;
            this.btnReload2.Location = new System.Drawing.Point(288, 28);
            this.btnReload2.Name = "btnReload2";
            this.btnReload2.Size = new System.Drawing.Size(75, 20);
            this.btnReload2.TabIndex = 19;
            this.btnReload2.Text = "Reload";
            this.ttpChecker.SetToolTip(this.btnReload2, "Reload the string array from file");
            this.btnReload2.Click += new System.EventHandler(this.btnReload2_Click);
            // 
            // btnReload1
            // 
            this.btnReload1.Enabled = false;
            this.btnReload1.Location = new System.Drawing.Point(4, 28);
            this.btnReload1.Name = "btnReload1";
            this.btnReload1.Size = new System.Drawing.Size(75, 20);
            this.btnReload1.TabIndex = 20;
            this.btnReload1.Text = "Reload";
            this.ttpChecker.SetToolTip(this.btnReload1, "Reload the string array from file");
            this.btnReload1.Click += new System.EventHandler(this.btnReload1_Click);
            // 
            // dgSource
            // 
            this.dgSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
                | System.Windows.Forms.AnchorStyles.Left)));
            this.dgSource.CaptionText = "GUID";
            this.dgSource.DataMember = "";
            this.dgSource.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dgSource.Location = new System.Drawing.Point(4, 164);
            this.dgSource.Name = "dgSource";
            this.dgSource.ReadOnly = true;
            this.dgSource.Size = new System.Drawing.Size(284, 216);
            this.dgSource.TabIndex = 2;
            // 
            // dgResult
            // 
            this.dgResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
                | System.Windows.Forms.AnchorStyles.Left) 
                | System.Windows.Forms.AnchorStyles.Right)));
            this.dgResult.CaptionText = "GUID";
            this.dgResult.DataMember = "";
            this.dgResult.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dgResult.Location = new System.Drawing.Point(292, 164);
            this.dgResult.Name = "dgResult";
            this.dgResult.ReadOnly = true;
            this.dgResult.Size = new System.Drawing.Size(284, 216);
            this.dgResult.TabIndex = 3;
            // 
            // toolBar1
            // 
            this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
                                                                                        this.tbrBtnSave1,
                                                                                        this.tbrBtnExport1});
            this.toolBar1.ButtonSize = new System.Drawing.Size(17, 17);
            this.toolBar1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new System.Drawing.Point(4, 136);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new System.Drawing.Size(48, 28);
            this.toolBar1.TabIndex = 17;
            this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
            // 
            // tbrBtnSave1
            // 
            this.tbrBtnSave1.Enabled = false;
            this.tbrBtnSave1.ImageIndex = 0;
            this.tbrBtnSave1.ToolTipText = "Save the current data grid to string arracy";
            // 
            // tbrBtnExport1
            // 
            this.tbrBtnExport1.Enabled = false;
            this.tbrBtnExport1.ImageIndex = 1;
            this.tbrBtnExport1.ToolTipText = "Export current data grid to text file";
            // 
            // imageList1
            // 
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // toolBar2
            // 
            this.toolBar2.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
                                                                                        this.tbrBtnSave2,
                                                                                        this.tbrBtnExport2});
            this.toolBar2.ButtonSize = new System.Drawing.Size(17, 17);
            this.toolBar2.Dock = System.Windows.Forms.DockStyle.None;
            this.toolBar2.DropDownArrows = true;
            this.toolBar2.ImageList = this.imageList1;
            this.toolBar2.Location = new System.Drawing.Point(296, 136);
            this.toolBar2.Name = "toolBar2";
            this.toolBar2.ShowToolTips = true;
            this.toolBar2.Size = new System.Drawing.Size(48, 28);
            this.toolBar2.TabIndex = 18;
            this.toolBar2.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar2_ButtonClick);
            // 
            // tbrBtnSave2
            // 
            this.tbrBtnSave2.Enabled = false;
            this.tbrBtnSave2.ImageIndex = 0;
            this.tbrBtnSave2.ToolTipText = "Save the current data grid to string arracy";
            // 
            // tbrBtnExport2
            // 
            this.tbrBtnExport2.Enabled = false;
            this.tbrBtnExport2.ImageIndex = 1;
            this.tbrBtnExport2.ToolTipText = "Export current data grid to text file";
            // 
            // CheckerWnd
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(580, 385);
            this.Controls.Add(this.btnReload1);
            this.Controls.Add(this.btnReload2);
            this.Controls.Add(this.toolBar2);
            this.Controls.Add(this.toolBar1);
            this.Controls.Add(this.btnRsuCompare);
            this.Controls.Add(this.btnSrcCompare);
            this.Controls.Add(this.btnRefreshRsu);
            this.Controls.Add(this.btnRefreshSrc);
            this.Controls.Add(this.btnDeDupRsu);
            this.Controls.Add(this.btnGetRsuDup);
            this.Controls.Add(this.btnGetSrcDup);
            this.Controls.Add(this.btnDeDupSrc);
            this.Controls.Add(this.rtbResult);
            this.Controls.Add(this.rtbSource);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.dgResult);
            this.Controls.Add(this.dgSource);
            this.Controls.Add(this.lnkResultFile);
            this.Controls.Add(this.lnkSourceFile);
            this.Name = "CheckerWnd";
            this.Text = "CheckerWnd";
            ((System.ComponentModel.ISupportInitialize)(this.dgSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

        /// <summary>
        /// Converts the contents of a text file into a string array
        /// Read each line of text file, and added to string array.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string[] FileToStrArray(string fileName)
        {
            ArrayList al = new ArrayList();
            StreamReader sr = File.OpenText(fileName);
            
            string str = sr.ReadLine(); // read the first line
            while( str != null )
            {
                al.Add(str);
                str = sr.ReadLine(); // read next line
            }
            sr.Close();

            // allocate the array size by counting the size of arraylist
            string[] strArray = new string[al.Count];
            al.CopyTo( strArray );

            return( strArray );
        }//end of FileToStrArray

        /// <summary>
        /// Saves the content of string array to a text file. 
        /// Loops through the array and writes each string in the array to the text file. 
        /// </summary>
        /// <param name="strArray">string array</param>
        /// <param name="fileName">Full path file name</param>
        public void StrArrayToFile(string[] strArray, string path_FileName)
        {
            StreamWriter sw = File.CreateText(path_FileName);
            foreach( string str in strArray )
            {
                sw.WriteLine(str.Trim('"'));
            }
            sw.Close();
        }//end of ArrayToFile

        /// <summary>
        /// Converts an array of strings to a DataTable
        /// This method can be used with any type of single character delimiter 
        /// and multiple single character delimiters, but does not ignore delimiters 
        /// if they are inside quotation marks.
        /// e.g. Input
        /// string[] list = {"Symbols","IBM","INTC","DELL","SUNW","MSFT"};
        /// DataTable dt = ArrayToDataTable(list, new char[1]{','}, true);
        /// Output: a column of "Symbols","IBM","INTC","DELL","SUNW","MSFT"
        /// </summary>
        /// <param name="strArray"></param>
        /// <param name="delimiter"></param>
        /// <param name="bHeaders"></param>
        /// <returns>DataTable - dt</returns>
        public DataTable ArrayToDataTable(string[] strArray, char[] delimiter, bool bHeaders)
        {
            // TO DO: better error handling
            if( 0x1000000 < strArray.Length ) // max row of data table = 16,777,216
                return null;

            DataTable dt = new DataTable();

            // assume the first array element contain header info
            string[] header = strArray[0].Split(delimiter); 
            foreach( string str in header )
            {
                if( bHeaders ) 
                {
                    dt.Columns.Add(str);
                }
                else
                {
                    dt.Columns.Add();
                }
            }//end of foreach

            for (int iRow = 0; iRow < strArray.Length; iRow++ )
            {
                if( !( bHeaders && (iRow == 0)) )
                {
                    string str = strArray[iRow];
                    string[] item = str.Split(delimiter);
                    DataRow dr = dt.NewRow();
                    for( int iCol=0; iCol < item.Length; iCol++)
                    {
                        dr[iCol] = item[iCol];
                    }
                    dt.Rows.Add(dr);
                }//end of if
            }// end of for

            return( dt );
        }//end of ArrayToDataTable

        /// <summary>
        /// This method is for using with strings in standard csv format. 
        /// If the commas are inside quotation marks, they are ignored. 
        /// Double quotation marks are considered as single quotation marks. 
        /// e.g. Input
        /// string path = @"C:\misc\houseLoan.csv";
        /// DataTable dt = ArrayToDataTable(FileToStrArray(path), true);
        /// Output: a table with nice columns and rows
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="bHeaders"></param>
        /// <returns>DataTable - dt</returns>
        public DataTable ArrayToDataTable(string[] arr, bool bHeaders)
        {
            // TO DO: error handling on input arr.

            DataTable dt = new DataTable();
            //pattern for matching csv format
            string pattern = "(?:^|,)(?:\"((?>[^\"]+|\"\")*)\"|([^\",]*))";
            Regex regex = new Regex(pattern);

            for( int iRow = 0; iRow < arr.Length; iRow++ )
            {
                Match match = regex.Match(arr[iRow]);
                DataRow dr = dt.NewRow();
                int iCol = 0;
                while( match.Success ) // matching reg expression
                {
                    string item = string.Empty;
                    if (match.Groups[1].Success)
                    {
                        item = match.Groups[1].Value.Replace("\"\"", "\"");
                    }
                    else
                    {
                        item = match.Groups[2].Value;
                    }
                    if( iRow == 0 )
                    {
                        if( bHeaders )
                        {
                            dt.Columns.Add(item);
                        }
                        else
                        {
                            dt.Columns.Add();
                            dr[iCol] = item;
                        }
                    }
                    else
                    {
                        dr[iCol] = item;
                    }
                    iCol++;
                    match = match.NextMatch();
                }//end of while - match.Success
                if( !((iRow == 0) && bHeaders) )
                {
                    dt.Rows.Add(dr);
                }
            }//end of for
            return( dt );
        }//end of ArrayToDataTable

        /// <summary>
        /// Browse the source GUID file and then load the content into DataGrid
        /// 1) Clear all varialbe.
        /// 2) Load the file...
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkSourceFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "CheckWnd.cs - lnkSourceFile_LinkClicked" );

            // initialize variables
            m_numMailSource = m_sourceDiffCnt = m_sourceDeDupCnt = 0; 

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.ShowReadOnly = true;
            ofDlg.RestoreDirectory = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtSource.Text = ofDlg.FileName;
            }//end of if
            else
            {
                Trace.WriteLine( "\t Cancel Button Hit - return" );
                return;
            }
            
            LoadFileToDataGrid( txtSource.Text, SOURCE_WND );
        }//end of lnkSourceFile_LinkClicked

        /// <summary>
        /// Browse the result GUID file and then load the content into DataGrid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lnkResultFile_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            Trace.WriteLine( "CheckWnd.cs - lnkResultFile_LinkClicked" );

            // initialize variables
            m_numMailResult = m_resultDiffCnt = m_resultDeDupCnt = 0;

            OpenFileDialog ofDlg = new OpenFileDialog();
            ofDlg.ShowReadOnly = true;
            ofDlg.RestoreDirectory = true;
            if( ofDlg.ShowDialog() == DialogResult.OK )
            {
                txtResult.Text = ofDlg.FileName;
            }//end of if
            else
            {
                Trace.WriteLine( "\t Cancel Button Hit - return" );
                return;
            }
			
            LoadFileToDataGrid( txtResult.Text, RESULT_WND );
        }//end of lnkResultFile_LinkClicked

        /// <summary>
        /// Convert the text file content into sorted string array list.
        /// Record the number of element in array.
        /// Then convert the string array to data table.
        /// Bind the data table into data grid.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="idxDataGrid"> 0 - source data grid, 1 - result data grid</param>
        public void LoadFileToDataGrid(string fileName, int idxDataGrid)
        {
            if( commObj.CheckFileInfo( fileName ) == 1 ) // error - file doesn't exist
            {
                switch( idxDataGrid )
                {
                    case SOURCE_WND:
                        rtbSource.Text = "File doesn't exist!!";
                        txtSource.BackColor = Color.Red;
                        txtSource.Focus();                        
                        break;
                    case RESULT_WND:
                        rtbResult.Text = "File doesn't exist!!";
                        txtResult.BackColor = Color.Red;
                        txtResult.Focus();
                        break;
                }//end of switch

                return;
            }//end of if

            string[] sortedArray = FileToStrArray( fileName );
            Array.Sort( sortedArray );

            DataTable dTable = new DataTable();
            dTable = ArrayToDataTable( sortedArray, false );
            switch( idxDataGrid )
            {
                case 0:// source data grid  
                    dgSource.DataSource = dTable;
                    m_numMailSource = sortedArray.Length;
                    m_strArrSource  = sortedArray;
                    EnableSourceControl();
                    break;
                case 1:// result data grid
                    dgResult.DataSource = dTable;
                    m_numMailResult = sortedArray.Length;
                    m_strArrResult  = sortedArray;
                    EnableResultControl();
                    break;
            }//end of switch
            UpdateDisplayWnd( idxDataGrid );
        }//end of LoadFileToDataGrid

        /// <summary>
        /// Update all info parameter and then display them into rich text box        
        /// </summary>
        /// <param name="option">0 - for source window, 1 - for result window</param>
        public void UpdateDisplayWnd(int option)
        {
            switch( option )
            {
                case 0:
//                    rtbSource.Text = "Num mails: " + m_numMailSource + "\r\n"
                    rtbSource.Text = "Num mails: " + m_strArrSource.Length + "\r\n"
                        + "Diff Count: " + m_sourceDiffCnt + "\r\n"
                        + "DeDups Count: " + m_sourceDeDupCnt + "\r\n"
                        + m_strMsg;
                    break;
                case 1:
//                    rtbResult.Text = "Num mails: " + m_numMailResult + "\r\n"
                    rtbResult.Text = "Num mails: " + m_strArrResult.Length + "\r\n"
                        + "Diff Count: " + m_resultDiffCnt + "\r\n"
                        + "DeDups Count: " + m_resultDeDupCnt + "\r\n"
                        + m_strMsg;
                    break;
            }//end of switch

            m_strMsg = ""; // reset error message after display it.
        }//end of UpdateDisplayWnd

        private void KeyPress_txtSource(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            Debug.WriteLine("CheckerWnd.cs - KeyPress_txtSource");
            switch( e.KeyChar )
            {
                case (char)13: // ENTER key
                    LoadFileToDataGrid( txtSource.Text, SOURCE_WND );
                    break;
            }//end of switch     
        }//end of KeyPress_txtSource

        private void KeyPress_txtResult(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            Debug.WriteLine("CheckerWnd.cs - KeyPress_txtResult");
            switch( e.KeyChar )
            {
                case (char)13: // ENTER key
                    LoadFileToDataGrid( txtResult.Text, RESULT_WND );
                    break;
            }//end of switch
        }//end of KeyPress_txtResult

        /// <summary>
        /// Compare two arrays and get a list of missing items
        /// The 1st string array contains more items than the 2nd Array list.
        /// </summary>
        /// <param name="strArrMoreItems">More items</param>
        /// <param name="alLessItems">Less items</param>
        /// <returns>ArrayList - missing Items</returns>
        public ArrayList GetMissingItems(string[] strArrMoreItems, ArrayList alLessItems)
        {
            ArrayList alMissItems = new ArrayList();
            for( int i = 0; i < strArrMoreItems.Length; i++ )
            {
                if( 0 < strArrMoreItems[i].Trim().Length ) //filter out empty strings
                {
                    if( !alLessItems.Contains(strArrMoreItems[i].Trim()) )
                    {
                        alMissItems.Add(strArrMoreItems[i]);
                    }//end of if
                }//end of if
            }//end of for
            return( alMissItems );
        }//end of GetMissingItems

        /// <summary>
        /// Compare two arrays and get a list of missing items
        /// The 1st string array contains More items than the 2nd string array.
        /// </summary>
        /// <param name="strArrMoreItems">More Items</param>
        /// <param name="strArrLessItems">Less Items</param>
        /// <returns>string array - missing Items</returns>
        public string[] GetMissItems(string[] strArrMoreItems, string[] strArrLessItems)
        {
            ArrayList alMissItems = new ArrayList();
            for( int i = 0; i < strArrMoreItems.Length; i++ )
            {                
                bool match = false;
                if( 0 < strArrMoreItems[i].Trim().Length ) //filter out empty strings
                {
                    for( int j = 0; j < strArrLessItems.Length; j++ )
                    {
                        if( strArrLessItems[j].Trim().Equals(strArrMoreItems[i].Trim()) )
                        {
                            match = true;
                            break;
                        }
                    }// end of for
                    if( !match )
                    {
                        alMissItems.Add(strArrMoreItems[i]);
                        commObj.LogToFile("\t Adding Missing item");
                    }
                }//end of if - outer
            }// end of for - outer
            string[] tmpStrArr = new string[alMissItems.Count];
            alMissItems.CopyTo(tmpStrArr);

            return tmpStrArr;
        }// end of GetMissItems

        /// <summary>
        /// Remove the duplicated items, and return a unique string array
        /// </summary>
        /// <param name="strArrItems"></param>
        /// <param name="bSort"></param>
        /// <returns></returns>
        public string[] RemoveDups(string[] strArrItems, bool bSort)
        {
            ArrayList noDups = new ArrayList();
            for( int i = 0; i < strArrItems.Length; i++ )
            {
                if( !noDups.Contains(strArrItems[i].Trim()) )
                {
                    noDups.Add(strArrItems[i].Trim());
                }
            }// end of for
            if( bSort )
                noDups.Sort();  //sorts list alphabetically

            string[] uniquestrArrItems = new String[noDups.Count];
            noDups.CopyTo(uniquestrArrItems);
            return uniquestrArrItems;
        }//end of RemoveDups

        public string[] ExtractDups(string[] strArrItems, bool bSort)
        {
            ArrayList alTmp = new ArrayList();
            ArrayList alDup = new ArrayList(); 

            for( int i = 0; i < strArrItems.Length; i++ )
            {
                if( !alTmp.Contains(strArrItems[i].Trim()) )
                {
                    alTmp.Add( strArrItems[i].Trim() );
                }
                else
                {
                    alDup.Add( strArrItems[i].Trim() );
                }
            }//end of for
            if( bSort )
                alDup.Sort();  //sorts list alphabetically

            string[] strArrDups = new String[alDup.Count];
            alDup.CopyTo(strArrDups);
            return( strArrDups );
        }// end of ExtractDups

        private void btnDeDupSrc_Click(object sender, System.EventArgs e)
        {
            string[] strArr = RemoveDups( m_strArrSource, false );
            m_sourceDeDupCnt = strArr.Length;

            DataTable dTable = new DataTable();
            dTable = ArrayToDataTable( strArr, false );
            dgSource.DataSource = dTable;

            UpdateDisplayWnd( SOURCE_WND );
        }// end of btnDeDupSrc_Click

        private void btnDeDupRsu_Click(object sender, System.EventArgs e)
        {
            string[] strArr = RemoveDups( m_strArrResult, false );
            m_resultDeDupCnt = strArr.Length;

            DataTable dTable = new DataTable();
            dTable = ArrayToDataTable( strArr, false );
            dgResult.DataSource = dTable;

            UpdateDisplayWnd( RESULT_WND );        
        }//end of btnDeDupRsu_Click

        private void btnGetSrcDup_Click(object sender, System.EventArgs e)
        {
            string[] strArr = ExtractDups( m_strArrSource, false );
            m_sourceDiffCnt = strArr.Length;

            DataTable dTable = new DataTable();
            dTable = ArrayToDataTable( strArr, false );
            dgSource.DataSource = dTable;

            UpdateDisplayWnd( SOURCE_WND );
        }// end of btnGetSrcDup_Click

        private void btnGetRsuDup_Click(object sender, System.EventArgs e)
        {
            string[] strArr = ExtractDups( m_strArrResult, false );
            m_resultDiffCnt = strArr.Length;

            DataTable dTable = new DataTable();
            dTable = ArrayToDataTable( strArr, false );
            dgResult.DataSource = dTable;  

            UpdateDisplayWnd( RESULT_WND );
        }// end of btnGetRsuDup_Click

        /// <summary>
        /// This method converts a DataTable into an array of strings where each row of the table 
        /// becomes a string in the array. This is useful for converting a DataTable to a comma separated file.
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="delimiter"></param>
        /// <param name="headers"></param>
        /// <returns></returns>
        public static string[] DataTableToArray(DataTable dt, string delimiter, bool headers)
        {
            int rowCnt = (headers) ? (dt.Rows.Count+1) : dt.Rows.Count;
            string[] arr = new string[rowCnt];

            if( headers )
            {
                //write column headings
                string colstr = string.Empty;
                for (int icol=0; icol<dt.Columns.Count; icol++)
                {
                    DataColumn dc = dt.Columns[icol];
                    if (icol == 0) 
                    {
                        colstr = dc.ColumnName;
                    }
                    else
                    {
                        colstr += delimiter + dc.ColumnName;
                    }
                }
                arr[0] = colstr;
            }// end of if - header

            //write table data
            for (int irow=0; irow<dt.Rows.Count; irow++)
            {
                DataRow dr = dt.Rows[irow];
                string str = string.Empty;
                for (int icol=0; icol<dt.Columns.Count; icol++)
                {
                    //if item contains delimiter, put quotes around item
                    string item = (dr[icol].ToString().IndexOf(delimiter) > -1) ? "\"" + dr[icol].ToString() + "\"" : dr[icol].ToString();
                    if (icol == 0) 
                    {
                        str = item;
                    }
                    else
                    {
                        str += delimiter + item;
                    }
                }
                int add = (headers) ? 1 : 0; //add row if headers is true
                arr[irow+add] = str;
            }// end of for
            return arr;
        }// end of DataTableToArray

        /// <summary>
        /// Refresh whatever in the string array into data grid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRefreshSrc_Click(object sender, System.EventArgs e)
        {
// comment 032205
//            DataTable dTable = new DataTable();
            dgSource.DataSource = ArrayToDataTable( m_strArrSource, false );
            m_numMailSource = m_strArrSource.Length;
            UpdateDisplayWnd( SOURCE_WND );
        }// end of btnRefreshSrc_Click

        /// <summary>
        /// Refresh whatever in the string array into data grid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRefreshRsu_Click(object sender, System.EventArgs e)
        {
// comment 032205
//            DataTable dTable = new DataTable();
            dgResult.DataSource = ArrayToDataTable( m_strArrResult, false );
            m_numMailResult = m_strArrResult.Length;
            UpdateDisplayWnd( RESULT_WND );
        }// end of btnRefreshRsu_Click

		/// <summary>
		/// Compare the source string array against the result string array.
		/// Display items not in result string array.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSrcCompare_Click(object sender, System.EventArgs e)
		{
			if( (m_strArrSource == null) || (m_strArrResult == null) )
				return;

			string[] strArray = GetMissItems( m_strArrSource, m_strArrResult );
			m_sourceDiffCnt = strArray.Length;
			DataTable dTable = new DataTable();
			dgSource.DataSource = ArrayToDataTable( strArray, false );
			UpdateDisplayWnd( SOURCE_WND );
		}// end of btnSrcCompare_Click

		/// <summary>
		/// Compare the result string array against the source string array.
		/// Display items NOT in source string array.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRsuCompare_Click(object sender, System.EventArgs e)
		{
			if( (m_strArrSource == null) || (m_strArrResult == null) )
				return;

			string [] strArray = GetMissItems( m_strArrResult, m_strArrSource );
			m_resultDiffCnt = strArray.Length;
			DataTable dTable = ArrayToDataTable( strArray, false );
			dgResult.DataSource = dTable;

			UpdateDisplayWnd( RESULT_WND );		
		}// end of btnRsuCompare_Click

        /// <summary>
        /// Toolbar event - check the index of toolbar button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			// Evaluate the Button property of the ToolBarButtonClickEventArgs
			// to determine which button was clicked.
			switch (toolBar1.Buttons.IndexOf(e.Button))
			{
				case 0 : // save the data grid to string array
                    SaveDataGridToStrArray( dgSource, ref m_strArrSource );
					break;
				case 1 :
                    SaveDataGridToStrArray( dgSource, ref m_strArrSource );
                    StrArrayToFile( m_strArrSource, GetSaveAbsPathFileName() );
					break;
			}// end of switch

            UpdateDisplayWnd( SOURCE_WND );
		}// end of toolBar1_ButtonClick

		private void toolBar2_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (toolBar2.Buttons.IndexOf(e.Button))
			{
				case 0 : // save the data grid to string array
                    SaveDataGridToStrArray( dgResult, ref m_strArrResult );
					break;
				case 1 :
					SaveDataGridToStrArray( dgResult, ref m_strArrResult );
                    StrArrayToFile( m_strArrResult, GetSaveAbsPathFileName() );
                    break;
			}//end of switch

	        UpdateDisplayWnd( RESULT_WND );
		}//end of toolBar2_ButtonClick

        public void SaveDataGridToStrArray( DataGrid dg, ref string[] strArray )
        {
            DataTable dTable = new DataTable();
            dTable = (DataTable)dg.DataSource;
            if( dTable == null )
            {
                m_strMsg = "Error - Data Grid Empty";
            }//end of if
            strArray = DataTableToArray( dTable, "", false );
        }// end of SaveDataGridToStrArray

        /// <summary>
        /// Reload from the file to data grid. NOT from strArray in memory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReload2_Click(object sender, System.EventArgs e)
        {
            LoadFileToDataGrid( txtResult.Text, RESULT_WND );
        }// end of btnReload2_Click

        /// <summary>
        /// Reload from the file to data grid. NOT from strArray in memory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReload1_Click(object sender, System.EventArgs e)
        {
            LoadFileToDataGrid( txtSource.Text, SOURCE_WND );
        }// end of btnReload1_Click

        public void EnableSourceControl()
        {
            btnDeDupSrc.Enabled     = true;
            btnGetSrcDup.Enabled    = true;
            btnRefreshSrc.Enabled   = true;
            btnSrcCompare.Enabled   = true;
            tbrBtnSave1.Enabled     = true;
            tbrBtnExport1.Enabled   = true;
            btnReload1.Enabled      = true;
        }// end of EnableSourceControl

        public void EnableResultControl()
        {
            btnDeDupRsu.Enabled     = true;
            btnGetRsuDup.Enabled    = true;
            btnRefreshRsu.Enabled   = true;
            btnRsuCompare.Enabled   = true;
            tbrBtnSave2.Enabled     = true;
            tbrBtnExport2.Enabled   = true;
            btnReload2.Enabled      = true;        
        }// end of EnableResultControl

        public string GetSaveAbsPathFileName()
        {
            string filename = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog(); 
            saveFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"  ;
            saveFileDialog.FilterIndex = 2 ;
            saveFileDialog.RestoreDirectory = true ;
 
            if(saveFileDialog.ShowDialog() == DialogResult.OK)
                filename = saveFileDialog.FileName;

            return( filename );
        }
	}//end of class - CheckerWnd
}//end of Name Space - QATool
