using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for ucFTPClient.
	/// </summary>
	public class ucFTPClient : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.LinkLabel lnkFolder;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.ComboBox cboFTP;
        private System.Windows.Forms.Label lblFTPServer;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.ComboBox cboUserId;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Label lblUserId;
        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.CheckBox chkDebug;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ToolTip ttpFTP;
        private System.ComponentModel.IContainer components;

		public ucFTPClient()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            lblStatus.Text = "";
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
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.lnkFolder = new System.Windows.Forms.LinkLabel();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.cboFTP = new System.Windows.Forms.ComboBox();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblFTPServer = new System.Windows.Forms.Label();
            this.ttpFTP = new System.Windows.Forms.ToolTip(this.components);
            this.cboUserId = new System.Windows.Forms.ComboBox();
            this.chkDebug = new System.Windows.Forms.CheckBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserId = new System.Windows.Forms.Label();
            this.btnTest = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtFolder
            // 
            this.txtFolder.Location = new System.Drawing.Point(60, 52);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(132, 20);
            this.txtFolder.TabIndex = 103;
            this.txtFolder.Text = "d:\\download";
            this.ttpFTP.SetToolTip(this.txtFolder, "Download location");
            // 
            // lnkFolder
            // 
            this.lnkFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.lnkFolder.Location = new System.Drawing.Point(8, 52);
            this.lnkFolder.Name = "lnkFolder";
            this.lnkFolder.Size = new System.Drawing.Size(48, 16);
            this.lnkFolder.TabIndex = 102;
            this.lnkFolder.TabStop = true;
            this.lnkFolder.Text = "Folder :";
            this.lnkFolder.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ttpFTP.SetToolTip(this.lnkFolder, "Download Folder");
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 13;
            this.cboPort.Location = new System.Drawing.Point(244, 28);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(44, 21);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 101;
            this.cboPort.Text = "21";
            this.ttpFTP.SetToolTip(this.cboPort, "FTP Server Port Number");
            // 
            // cboFTP
            // 
            this.cboFTP.ItemHeight = 13;
            this.cboFTP.Items.AddRange(new object[] {
                                                        "10.1.21.191",
                                                        "10.1.42.201",
                                                        "10.1.42.203"});
            this.cboFTP.Location = new System.Drawing.Point(60, 28);
            this.cboFTP.Name = "cboFTP";
            this.cboFTP.Size = new System.Drawing.Size(132, 21);
            this.cboFTP.Sorted = true;
            this.cboFTP.TabIndex = 100;
            this.cboFTP.Text = "10.1.42.201";
            this.ttpFTP.SetToolTip(this.cboFTP, "FTP Server IP");
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(196, 32);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(38, 16);
            this.lblPort.TabIndex = 99;
            this.lblPort.Text = "Port # ";
            // 
            // lblFTPServer
            // 
            this.lblFTPServer.Location = new System.Drawing.Point(0, 32);
            this.lblFTPServer.Name = "lblFTPServer";
            this.lblFTPServer.Size = new System.Drawing.Size(64, 16);
            this.lblFTPServer.TabIndex = 98;
            this.lblFTPServer.Text = "FTP Server";
            // 
            // cboUserId
            // 
            this.cboUserId.Items.AddRange(new object[] {
                                                           ""});
            this.cboUserId.Location = new System.Drawing.Point(60, 4);
            this.cboUserId.Name = "cboUserId";
            this.cboUserId.Size = new System.Drawing.Size(132, 21);
            this.cboUserId.TabIndex = 106;
            this.cboUserId.Text = "kent";
            this.ttpFTP.SetToolTip(this.cboUserId, "FTP user name");
            // 
            // chkDebug
            // 
            this.chkDebug.Location = new System.Drawing.Point(292, 32);
            this.chkDebug.Name = "chkDebug";
            this.chkDebug.Size = new System.Drawing.Size(64, 16);
            this.chkDebug.TabIndex = 109;
            this.chkDebug.Text = "Debug";
            this.ttpFTP.SetToolTip(this.chkDebug, "Debug log for ftp connection");
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(244, 4);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(96, 20);
            this.txtPassword.TabIndex = 107;
            this.txtPassword.Text = "skyline";
            // 
            // lblPassword
            // 
            this.lblPassword.Location = new System.Drawing.Point(192, 8);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(56, 16);
            this.lblPassword.TabIndex = 105;
            this.lblPassword.Text = "Password";
            this.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblUserId
            // 
            this.lblUserId.Location = new System.Drawing.Point(4, 8);
            this.lblUserId.Name = "lblUserId";
            this.lblUserId.Size = new System.Drawing.Size(52, 16);
            this.lblUserId.TabIndex = 104;
            this.lblUserId.Text = "User ID";
            this.lblUserId.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(344, 4);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(40, 20);
            this.btnTest.TabIndex = 108;
            this.btnTest.Text = "Test";
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblStatus.Location = new System.Drawing.Point(4, 404);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(380, 20);
            this.lblStatus.TabIndex = 110;
            // 
            // ucFTPClient
            // 
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.chkDebug);
            this.Controls.Add(this.btnTest);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.cboUserId);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUserId);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.lnkFolder);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.cboFTP);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblFTPServer);
            this.Name = "ucFTPClient";
            this.Size = new System.Drawing.Size(388, 428);
            this.ResumeLayout(false);

        }
		#endregion

        private void btnTest_Click(object sender, System.EventArgs e)
        {
            Trace.WriteLine("ucFTPClient.cs - btnTest_Click");
            Debug.WriteLine("\t FTP: " + cboFTP.Text + " UID: " + cboUserId.Text + " Password: " + txtPassword.Text );

            FtpClient ftp = new FtpClient( cboFTP.Text, cboUserId.Text, txtPassword.Text );

            if( chkDebug.Checked )
                ftp.VerboseDebugging = true;

            ftp.Login();
            ftp.Upload(CreateFtpTestFile());
            ftp.Download("Upload.txt", txtFolder.Text + "\\download.txt");

            lblStatus.Text = ftp.Message;

            ftp.Close();
        }//end of btnTest_Click

        private string CreateFtpTestFile()
        {
            string fpfn = txtFolder.Text + "\\Upload.txt"; //fpfn - Full Path File Name
            DirectoryInfo di = new DirectoryInfo( txtFolder.Text );

            try
            {
                if( !di.Exists ) // directory not exist
                    di.Create(); // create one

                if( !File.Exists( fpfn ) )
                {   // Create the file.
                    using( FileStream fs = File.Create(fpfn, 1024) )
                    {
                        Byte[] info = new UTF8Encoding(true).GetBytes("This is some text in the upload file.");
                        // Add some information to the file.
                        fs.Write(info, 0, info.Length);
                    }//end of using
                }//end of if
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch

            return( fpfn );
        }//end of CreateFtpTestFile
	}
}
