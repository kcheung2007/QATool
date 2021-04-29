using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;


namespace QATool
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class ToolBox : System.Windows.Forms.Form
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;	
		private System.Windows.Forms.TabControl tabCtrl;
		private System.Windows.Forms.TabPage    tabMailPage;
		private System.Windows.Forms.TabPage    tabBatchMail;
        private System.Windows.Forms.TabPage    tabZelMsg;
        private System.Windows.Forms.TabPage    tabPstCount;
		private System.Windows.Forms.TabPage    tabDBMail;
		private System.Windows.Forms.TabPage	tabTemplate;
		private System.Windows.Forms.TabPage    tabGuiPage;
		private System.Windows.Forms.TabPage	tabACPage;
		private System.Windows.Forms.TabPage	tabOLPage;
        private System.Windows.Forms.TabPage	tabLNPage;
        private System.Windows.Forms.TabPage    tabNsfCount;
        private System.Windows.Forms.TabPage    tabFtpPage;
        
		// custom setting		
		private QATool.MailPage    mailPage;			
		private QATool.BatMailPage batchPage;
		private QATool.DBMailPage  dbPage;
        private QATool.ZelMsgPage  zelPage;
        private QATool.PstCounter  pstPage;
        private QATool.NSFCounter  nsfPage;
		private QATool.Template    tempPage;
		private QATool.GUITestPage guiPage;
		private QATool.ACTestPage  acPage;
		private QATool.OutlookPage olPage; // outlook page
        private QATool.NotesClient lnPage; // lotus notes page
        private QATool.ucFTPClient ftpPage;        
                
        static TraceSwitch traceSwitch = new TraceSwitch("MyTrace", "Trace Switch Demo");
        private System.Windows.Forms.TabPage tabMrvpPage;
        private QATool.ucMRVP ucMRVP1;
                				
		private QATool.CommObj   commObj = new CommObj();

		public ToolBox()
		{
            Trace.WriteLineIf( traceSwitch.TraceInfo, "Test Trace Info", traceSwitch.Description );
			// Required for Windows Form Designer support
			InitializeComponent(); // Define Tab control items order

            commObj.LogToFile("Create MRVP Page");
            ftpPage = new ucFTPClient(); // Create FTP Page
            ftpPage.Location = new Point(0,0);
            this.tabFtpPage.Controls.Add( ftpPage );

			commObj.LogToFile("Create Mail Page");
			mailPage = new MailPage(); // Create Mail Page
			mailPage.Location = new Point(0,0);
			this.tabMailPage.Controls.Add( mailPage );

			commObj.LogToFile("Create Batch Mail Page");
			batchPage = new BatMailPage(); // Create Batch Mail Page
			batchPage.Location = new Point(0,0);
			this.tabBatchMail.Controls.Add( batchPage );

            commObj.LogToFile("Create ZEL Message Mail Page");
            zelPage = new ZelMsgPage(); // Create Batch Mail Page
            zelPage.Location = new Point(0,0);
            this.tabZelMsg.Controls.Add( zelPage );

            commObj.LogToFile("Create PST Counter Page");
            pstPage = new PstCounter(); // Create Batch Mail Page
            pstPage.Location = new Point(0,0);
            this.tabPstCount.Controls.Add( pstPage );
            
            commObj.LogToFile("Create NSF Counter Page");
            nsfPage = new NSFCounter(); // Create Batch Mail Page
            nsfPage.Location = new Point(0,0);
            this.tabNsfCount.Controls.Add( nsfPage );

            commObj.LogToFile("Create DB Mail Page");
			dbPage = new DBMailPage(); // Create Batch Mail Page
			dbPage.Location = new Point(0,0);
			this.tabDBMail.Controls.Add( dbPage );

			commObj.LogToFile("Create GUI Test Page");
			guiPage = new GUITestPage();
			guiPage.Location = new Point(0,0);
			this.tabGuiPage.Controls.Add( guiPage );

			commObj.LogToFile("Create Audit Template Page");
			tempPage = new Template();
			tempPage.Location = new Point(0,0);
			this.tabTemplate.Controls.Add( tempPage );

			commObj.LogToFile("Create Audit Center Test Page");
			acPage = new ACTestPage();
			acPage.Location = new Point(0,0);
            acPage.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabACPage.Controls.Add( acPage );

			commObj.LogToFile("Create Outlook Page");
			olPage = new OutlookPage();
			olPage.Location = new Point(0,0);
            olPage.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabOLPage.Controls.Add( olPage );

            commObj.LogToFile("Create Lotus Notes Page");
            lnPage = new NotesClient();
            lnPage.Location = new Point(0,0);
            lnPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabLNPage.Controls.Add( lnPage );

		}// end of constructor

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			commObj.LogToFile("End of QATool - Execute Dispose");
			if( disposing )
			{
				if (components != null) 
				{
					Debug.WriteLine("\t Dispose component");
                    commObj.LogToFile("\t - clean up component");
					components.Dispose();
				}
			}
			base.Dispose( disposing );
            commObj.LogToFile("\t - finished disposing: three steps");
		}//end of Dispose

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.tabCtrl = new System.Windows.Forms.TabControl();
            this.tabMailPage = new System.Windows.Forms.TabPage();
            this.tabDBMail = new System.Windows.Forms.TabPage();
            this.tabACPage = new System.Windows.Forms.TabPage();
            this.tabLNPage = new System.Windows.Forms.TabPage();
            this.tabPstCount = new System.Windows.Forms.TabPage();
            this.tabZelMsg = new System.Windows.Forms.TabPage();
            this.tabOLPage = new System.Windows.Forms.TabPage();
            this.tabTemplate = new System.Windows.Forms.TabPage();
            this.tabBatchMail = new System.Windows.Forms.TabPage();
            this.tabGuiPage = new System.Windows.Forms.TabPage();
            this.tabNsfCount = new System.Windows.Forms.TabPage();
            this.tabFtpPage = new System.Windows.Forms.TabPage();
            this.tabMrvpPage = new System.Windows.Forms.TabPage();
            this.ucMRVP1 = new QATool.ucMRVP();
            this.tabCtrl.SuspendLayout();
            this.tabMrvpPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabCtrl
            // 
            this.tabCtrl.Controls.Add(this.tabMailPage);
            this.tabCtrl.Controls.Add(this.tabTemplate);
            this.tabCtrl.Controls.Add(this.tabOLPage);
            this.tabCtrl.Controls.Add(this.tabBatchMail);
            this.tabCtrl.Controls.Add(this.tabDBMail);
            this.tabCtrl.Controls.Add(this.tabACPage);
            this.tabCtrl.Controls.Add(this.tabLNPage);
            this.tabCtrl.Controls.Add(this.tabPstCount);
            this.tabCtrl.Controls.Add(this.tabZelMsg);
            this.tabCtrl.Controls.Add(this.tabGuiPage);
            this.tabCtrl.Controls.Add(this.tabNsfCount);
            this.tabCtrl.Controls.Add(this.tabFtpPage);
            this.tabCtrl.Controls.Add(this.tabMrvpPage);
            this.tabCtrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabCtrl.Location = new System.Drawing.Point(0, 0);
            this.tabCtrl.Multiline = true;
            this.tabCtrl.Name = "tabCtrl";
            this.tabCtrl.SelectedIndex = 0;
            this.tabCtrl.Size = new System.Drawing.Size(394, 491);
            this.tabCtrl.TabIndex = 0;
            this.tabCtrl.SelectedIndexChanged += new System.EventHandler(this.tabCtrl_SelectedIndexChanged);
            // 
            // tabMailPage
            // 
            this.tabMailPage.Location = new System.Drawing.Point(4, 58);
            this.tabMailPage.Name = "tabMailPage";
            this.tabMailPage.Size = new System.Drawing.Size(386, 429);
            this.tabMailPage.TabIndex = 0;
            this.tabMailPage.Text = "Mail Page";
            // 
            // tabDBMail
            // 
            this.tabDBMail.Location = new System.Drawing.Point(4, 40);
            this.tabDBMail.Name = "tabDBMail";
            this.tabDBMail.Size = new System.Drawing.Size(402, 427);
            this.tabDBMail.TabIndex = 2;
            this.tabDBMail.Text = "DB Mail";
            // 
            // tabACPage
            // 
            this.tabACPage.Location = new System.Drawing.Point(4, 40);
            this.tabACPage.Name = "tabACPage";
            this.tabACPage.Size = new System.Drawing.Size(402, 427);
            this.tabACPage.TabIndex = 4;
            this.tabACPage.Text = "AC Check";
            // 
            // tabLNPage
            // 
            this.tabLNPage.Location = new System.Drawing.Point(4, 40);
            this.tabLNPage.Name = "tabLNPage";
            this.tabLNPage.Size = new System.Drawing.Size(402, 427);
            this.tabLNPage.TabIndex = 10;
            this.tabLNPage.Text = "Notes Client";
            // 
            // tabPstCount
            // 
            this.tabPstCount.Location = new System.Drawing.Point(4, 40);
            this.tabPstCount.Name = "tabPstCount";
            this.tabPstCount.Size = new System.Drawing.Size(402, 427);
            this.tabPstCount.TabIndex = 8;
            this.tabPstCount.Text = "PST Counter";
            // 
            // tabZelMsg
            // 
            this.tabZelMsg.Location = new System.Drawing.Point(4, 40);
            this.tabZelMsg.Name = "tabZelMsg";
            this.tabZelMsg.Size = new System.Drawing.Size(402, 427);
            this.tabZelMsg.TabIndex = 7;
            this.tabZelMsg.Text = "ZEL Msg";
            // 
            // tabOLPage
            // 
            this.tabOLPage.Location = new System.Drawing.Point(4, 58);
            this.tabOLPage.Name = "tabOLPage";
            this.tabOLPage.Size = new System.Drawing.Size(386, 429);
            this.tabOLPage.TabIndex = 5;
            this.tabOLPage.Text = "OutLook";
            // 
            // tabTemplate
            // 
            this.tabTemplate.Location = new System.Drawing.Point(4, 58);
            this.tabTemplate.Name = "tabTemplate";
            this.tabTemplate.Size = new System.Drawing.Size(386, 429);
            this.tabTemplate.TabIndex = 6;
            this.tabTemplate.Text = "Template";
            // 
            // tabBatchMail
            // 
            this.tabBatchMail.Location = new System.Drawing.Point(4, 58);
            this.tabBatchMail.Name = "tabBatchMail";
            this.tabBatchMail.Size = new System.Drawing.Size(386, 429);
            this.tabBatchMail.TabIndex = 3;
            this.tabBatchMail.Text = "Batch Mail";
            // 
            // tabGuiPage
            // 
            this.tabGuiPage.Location = new System.Drawing.Point(4, 40);
            this.tabGuiPage.Name = "tabGuiPage";
            this.tabGuiPage.Size = new System.Drawing.Size(402, 427);
            this.tabGuiPage.TabIndex = 1;
            this.tabGuiPage.Text = "GUI Test";
            // 
            // tabNsfCount
            // 
            this.tabNsfCount.Location = new System.Drawing.Point(4, 40);
            this.tabNsfCount.Name = "tabNsfCount";
            this.tabNsfCount.Size = new System.Drawing.Size(402, 427);
            this.tabNsfCount.TabIndex = 9;
            this.tabNsfCount.Text = "NSF Count";
            // 
            // tabFtpPage
            // 
            this.tabFtpPage.Location = new System.Drawing.Point(4, 40);
            this.tabFtpPage.Name = "tabFtpPage";
            this.tabFtpPage.Size = new System.Drawing.Size(402, 427);
            this.tabFtpPage.TabIndex = 11;
            this.tabFtpPage.Text = "FTP Page";
            // 
            // tabMrvpPage
            // 
            this.tabMrvpPage.Controls.Add(this.ucMRVP1);
            this.tabMrvpPage.Location = new System.Drawing.Point(4, 58);
            this.tabMrvpPage.Name = "tabMrvpPage";
            this.tabMrvpPage.Size = new System.Drawing.Size(386, 429);
            this.tabMrvpPage.TabIndex = 12;
            this.tabMrvpPage.Text = "MRVP";
            this.tabMrvpPage.ToolTipText = "MRVP Project";
            // 
            // ucMRVP1
            // 
            this.ucMRVP1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ucMRVP1.Location = new System.Drawing.Point(0, 0);
            this.ucMRVP1.Name = "ucMRVP1";
            this.ucMRVP1.Size = new System.Drawing.Size(386, 429);
            this.ucMRVP1.TabIndex = 0;
            // 
            // ToolBox
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(394, 491);
            this.Controls.Add(this.tabCtrl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ToolBox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tool Box";
            this.Load += new System.EventHandler(this.ToolBox_Load);
            this.Closed += new System.EventHandler(this.ToolBox_Closed);
            this.tabCtrl.ResumeLayout(false);
            this.tabMrvpPage.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new ToolBox());
		}// end of Main

		private void ToolBox_Load(object sender, System.EventArgs e)
		{
		    this.Text = Application.ExecutablePath;
		}//end of ToolBox_Load

		private void tabCtrl_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Trace.WriteLine("MainForm.cs - tabCtrl_SelectedIndexChanged");
			Debug.WriteLine( tabCtrl.SelectedTab.ToString() );

			switch( tabCtrl.SelectedIndex ) // kill the instance here ??
			{
				case 0: // File Page
					break;
				case 1: // Directory Page
				{
				}
					break;
				case 2: // ADO Page
				{
				}
					break;
				case 3: // Control Page
					break;
			}// end of switch - Selected Index		
		}//end of tabCtrl_SelectedIndexChanged

        /// <summary>
        /// Close TCP connection????
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolBox_Closed(object sender, System.EventArgs e)
        {
            Trace.WriteLine("MainForm.cs - Main Application Exit...");
            commObj.LogToFile("MainForm.cs - ToolBox_Closed... Application Exit...");
            Application.Exit();
        }
	}
}
