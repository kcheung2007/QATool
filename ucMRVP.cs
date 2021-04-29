using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

using MySql.Data.MySqlClient;

namespace QATool
{
	/// <summary>
	/// Summary description for ucMRVP.
	/// </summary>
	public class ucMRVP : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.ComboBox cboUserId;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Label lblUserId;
        private System.Windows.Forms.ComboBox cboPort;
        private System.Windows.Forms.Label lblPort;
        private System.Windows.Forms.Label lblMySqlServer;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblDataBase;
        private System.Windows.Forms.ComboBox cboDatabase;
        private System.Windows.Forms.DataGrid dataGrid1;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.ToolTip ttpMrvp;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ComboBox cboMySqlIP;
        private System.Windows.Forms.ListBox lstTable;
        private System.Windows.Forms.ContextMenu dgContextMenu;
        private System.Windows.Forms.Button btnExpID;
        private System.Windows.Forms.Button btnGenGap; // data table of mySql

        private MySqlConnection     m_mySqlConn; // establish and always on? When to disconnect?
        private MySqlDataAdapter    m_mySqlAdapter;
        private MySqlCommandBuilder m_mySqlCmdBuilder;
        private DataTable           m_dtDataTable;        
        private string[]            m_strArray;

        private QATool.CommObj    commObj = new CommObj();

		public ucMRVP()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            commObj.InitComboBoxItem( cboPort, "[Port]" );
            commObj.InitComboBoxItem( cboMySqlIP, "[SQL IP]" );
            commObj.InitComboBoxItem( cboUserId, "[Login ID]" );

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
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ucMRVP));
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.cboUserId = new System.Windows.Forms.ComboBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserId = new System.Windows.Forms.Label();
            this.cboPort = new System.Windows.Forms.ComboBox();
            this.cboMySqlIP = new System.Windows.Forms.ComboBox();
            this.lblPort = new System.Windows.Forms.Label();
            this.lblMySqlServer = new System.Windows.Forms.Label();
            this.btnConnect = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblDataBase = new System.Windows.Forms.Label();
            this.cboDatabase = new System.Windows.Forms.ComboBox();
            this.lstTable = new System.Windows.Forms.ListBox();
            this.dataGrid1 = new System.Windows.Forms.DataGrid();
            this.dgContextMenu = new System.Windows.Forms.ContextMenu();
            this.lblTableName = new System.Windows.Forms.Label();
            this.ttpMrvp = new System.Windows.Forms.ToolTip(this.components);
            this.btnExpID = new System.Windows.Forms.Button();
            this.btnGenGap = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(264, 12);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(116, 20);
            this.txtPassword.TabIndex = 115;
            this.txtPassword.Text = "skyline";
            // 
            // cboUserId
            // 
            this.cboUserId.Items.AddRange(new object[] {
                                                           "root",
                                                           "kent"});
            this.cboUserId.Location = new System.Drawing.Point(80, 12);
            this.cboUserId.Name = "cboUserId";
            this.cboUserId.Size = new System.Drawing.Size(132, 21);
            this.cboUserId.TabIndex = 114;
            this.cboUserId.Text = "root";
            // 
            // lblPassword
            // 
            this.lblPassword.Location = new System.Drawing.Point(212, 16);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(56, 16);
            this.lblPassword.TabIndex = 113;
            this.lblPassword.Text = "Password";
            this.lblPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblUserId
            // 
            this.lblUserId.Location = new System.Drawing.Point(28, 16);
            this.lblUserId.Name = "lblUserId";
            this.lblUserId.Size = new System.Drawing.Size(52, 16);
            this.lblUserId.TabIndex = 112;
            this.lblUserId.Text = "User ID";
            this.lblUserId.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboPort
            // 
            this.cboPort.ItemHeight = 13;
            this.cboPort.Location = new System.Drawing.Point(264, 36);
            this.cboPort.Name = "cboPort";
            this.cboPort.Size = new System.Drawing.Size(52, 21);
            this.cboPort.Sorted = true;
            this.cboPort.TabIndex = 111;
            this.cboPort.Text = "3306";
            // 
            // cboMySqlIP
            // 
            this.cboMySqlIP.ItemHeight = 13;
            this.cboMySqlIP.Items.AddRange(new object[] {
                                                            "10.1.21.191",
                                                            "10.1.42.201",
                                                            "10.1.42.203"});
            this.cboMySqlIP.Location = new System.Drawing.Point(80, 36);
            this.cboMySqlIP.Name = "cboMySqlIP";
            this.cboMySqlIP.Size = new System.Drawing.Size(132, 21);
            this.cboMySqlIP.Sorted = true;
            this.cboMySqlIP.TabIndex = 110;
            this.cboMySqlIP.Text = "10.1.42.201";
            // 
            // lblPort
            // 
            this.lblPort.Location = new System.Drawing.Point(216, 40);
            this.lblPort.Name = "lblPort";
            this.lblPort.Size = new System.Drawing.Size(38, 16);
            this.lblPort.TabIndex = 109;
            this.lblPort.Text = "Port # ";
            // 
            // lblMySqlServer
            // 
            this.lblMySqlServer.Location = new System.Drawing.Point(8, 40);
            this.lblMySqlServer.Name = "lblMySqlServer";
            this.lblMySqlServer.Size = new System.Drawing.Size(72, 16);
            this.lblMySqlServer.TabIndex = 108;
            this.lblMySqlServer.Text = "MySql Server";
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(320, 36);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(56, 20);
            this.btnConnect.TabIndex = 116;
            this.btnConnect.Text = "Connect";
            this.ttpMrvp.SetToolTip(this.btnConnect, "Connect to DB Server");
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(4, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(380, 64);
            this.groupBox1.TabIndex = 118;
            this.groupBox1.TabStop = false;
            // 
            // lblDataBase
            // 
            this.lblDataBase.Location = new System.Drawing.Point(16, 72);
            this.lblDataBase.Name = "lblDataBase";
            this.lblDataBase.Size = new System.Drawing.Size(60, 16);
            this.lblDataBase.TabIndex = 119;
            this.lblDataBase.Text = "Database ";
            this.lblDataBase.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cboDatabase
            // 
            this.cboDatabase.ItemHeight = 13;
            this.cboDatabase.Location = new System.Drawing.Point(80, 68);
            this.cboDatabase.Name = "cboDatabase";
            this.cboDatabase.Size = new System.Drawing.Size(132, 21);
            this.cboDatabase.Sorted = true;
            this.cboDatabase.TabIndex = 120;
            this.cboDatabase.Text = "test";
            this.ttpMrvp.SetToolTip(this.cboDatabase, "Type in Default DB");
            this.cboDatabase.SelectedIndexChanged += new System.EventHandler(this.cboDatabase_SelectedIndexChanged);
            // 
            // lstTable
            // 
            this.lstTable.Location = new System.Drawing.Point(4, 120);
            this.lstTable.Name = "lstTable";
            this.lstTable.Size = new System.Drawing.Size(100, 303);
            this.lstTable.TabIndex = 121;
            this.lstTable.SelectedIndexChanged += new System.EventHandler(this.lstTable_SelectedIndexChanged);
            // 
            // dataGrid1
            // 
            this.dataGrid1.AlternatingBackColor = System.Drawing.Color.Gainsboro;
            this.dataGrid1.BackColor = System.Drawing.Color.Silver;
            this.dataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.dataGrid1.CaptionBackColor = System.Drawing.Color.DarkSlateBlue;
            this.dataGrid1.CaptionFont = new System.Drawing.Font("Tahoma", 8F);
            this.dataGrid1.CaptionForeColor = System.Drawing.Color.White;
            this.dataGrid1.ContextMenu = this.dgContextMenu;
            this.dataGrid1.DataMember = "";
            this.dataGrid1.FlatMode = true;
            this.dataGrid1.ForeColor = System.Drawing.Color.Black;
            this.dataGrid1.GridLineColor = System.Drawing.Color.White;
            this.dataGrid1.HeaderBackColor = System.Drawing.Color.DarkGray;
            this.dataGrid1.HeaderForeColor = System.Drawing.Color.Black;
            this.dataGrid1.LinkColor = System.Drawing.Color.DarkSlateBlue;
            this.dataGrid1.Location = new System.Drawing.Point(108, 96);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.ParentRowsBackColor = System.Drawing.Color.Black;
            this.dataGrid1.ParentRowsForeColor = System.Drawing.Color.White;
            this.dataGrid1.SelectionBackColor = System.Drawing.Color.DarkSlateBlue;
            this.dataGrid1.SelectionForeColor = System.Drawing.Color.White;
            this.dataGrid1.Size = new System.Drawing.Size(272, 328);
            this.dataGrid1.TabIndex = 122;
            // 
            // lblTableName
            // 
            this.lblTableName.Location = new System.Drawing.Point(8, 100);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(92, 16);
            this.lblTableName.TabIndex = 123;
            this.lblTableName.Text = "Table Names";
            // 
            // btnExpID
            // 
            this.btnExpID.Enabled = false;
            this.btnExpID.Image = ((System.Drawing.Image)(resources.GetObject("btnExpID.Image")));
            this.btnExpID.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btnExpID.Location = new System.Drawing.Point(216, 68);
            this.btnExpID.Name = "btnExpID";
            this.btnExpID.Size = new System.Drawing.Size(24, 23);
            this.btnExpID.TabIndex = 124;
            this.ttpMrvp.SetToolTip(this.btnExpID, "Export Seq ID/server name");
            this.btnExpID.Click += new System.EventHandler(this.btnExpID_Click);
            // 
            // btnGenGap
            // 
            this.btnGenGap.Image = ((System.Drawing.Image)(resources.GetObject("btnGenGap.Image")));
            this.btnGenGap.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btnGenGap.Location = new System.Drawing.Point(244, 68);
            this.btnGenGap.Name = "btnGenGap";
            this.btnGenGap.Size = new System.Drawing.Size(24, 23);
            this.btnGenGap.TabIndex = 125;
            this.ttpMrvp.SetToolTip(this.btnGenGap, "Process Milter Files");
            this.btnGenGap.Click += new System.EventHandler(this.btnGenGap_Click);
            // 
            // ucMRVP
            // 
            this.Controls.Add(this.btnGenGap);
            this.Controls.Add(this.btnExpID);
            this.Controls.Add(this.lblTableName);
            this.Controls.Add(this.dataGrid1);
            this.Controls.Add(this.lstTable);
            this.Controls.Add(this.cboDatabase);
            this.Controls.Add(this.lblDataBase);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.cboUserId);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUserId);
            this.Controls.Add(this.cboPort);
            this.Controls.Add(this.cboMySqlIP);
            this.Controls.Add(this.lblPort);
            this.Controls.Add(this.lblMySqlServer);
            this.Controls.Add(this.groupBox1);
            this.Name = "ucMRVP";
            this.Size = new System.Drawing.Size(388, 428);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
            this.ResumeLayout(false);

        }
		#endregion

        /// <summary>
        /// Build the Connection string and connect to the mysql server. 
        /// A default database is required for the connection. Most of the time it is "mysql".
        /// In this case, the database is input by user. Then will populate the DB table into list box.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConnect_Click(object sender, System.EventArgs e)
        {
            if( m_mySqlConn != null )
                m_mySqlConn.Close();

            string connStr = String.Format( "server={0}; user id={1}; password={2}; database={3}; pooling=false",
                                            cboMySqlIP.Text, cboUserId.Text, txtPassword.Text, cboDatabase.Text );

            try
            {
                m_mySqlConn = new MySqlConnection( connStr );
                m_mySqlConn.Open();

                GetMySqlDBList(); // get a list of mysql databases
                GetDBTableList( cboDatabase.Text ); // get a list of tables and populate the list box

            }//end of try
            catch( MySqlException mySqlEx )
            {
                MessageBox.Show( "Error connecting to the MySql Server " + mySqlEx.Message );
                Debug.WriteLine( mySqlEx.Message + "\n" + mySqlEx.GetType().ToString() + mySqlEx.StackTrace );
            }//end of catch
        }//end of btnConnect_Click

        /// <summary>
        /// One MySql server may contain multiple DB. Assume MySql already connected with a default DB.
        /// Then issue "SHOW DATABASES" command.
        /// </summary>
        private void GetMySqlDBList() 
        {
            Trace.WriteLine( "ucMRVP.cs - GetMySqlDBList()" );
            MySqlDataReader reader = null;

            MySqlCommand mySqlCmd = new MySqlCommand("SHOW DATABASES", m_mySqlConn);
            try 
            {
                reader = mySqlCmd.ExecuteReader();
                cboDatabase.Items.Clear();
                while( reader.Read() ) 
                {
                    cboDatabase.Items.Add( reader.GetString(0) );
                }
            }//end of try
            catch (MySqlException ex) 
            {
                MessageBox.Show("Failed to populate database list: " + ex.Message );
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
            finally 
            {
                if( reader != null )
                    reader.Close();
            }//end of finally
        }//end of GetMySqlDBList

        /// <summary>
        /// Get a list of tables for particular DB.
        /// 1) Clear the list box item
        /// 2) Clear the data table variable which is used for data binding the data grid
        /// 3) Clear the context menu which is used for exporting function
        /// 4) Add the table name into list box/// </summary>
        /// <param name="dbName">string: DB name</param>
        private void GetDBTableList( string dbName )
        {
            Trace.WriteLine("ucMRVP.cs - GetDBTableList");

            MySqlDataReader reader = null;
            m_mySqlConn.ChangeDatabase( dbName );

            MySqlCommand cmd = new MySqlCommand("SHOW TABLES", m_mySqlConn);
            try 
            {
                lstTable.Items.Clear();     // clear list table item
                if( m_dtDataTable != null ) // clear data table variable -> clear the data grid
                    m_dtDataTable.Clear();
                dataGrid1.CaptionText = ""; // clear data grid caption
                dgContextMenu.MenuItems.Clear(); // clear all items in context menu
                
                reader = cmd.ExecuteReader();                
                while( reader.Read() )
                {
                    lstTable.Items.Add( reader.GetString(0) ); // why 0?
                }
            }//end of try
            catch (MySqlException ex) 
            {
                MessageBox.Show("Failed to populate table list: " + ex.Message );
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
            finally 
            {
                if(reader != null) 
                    reader.Close();
            }//end of finally
        }//end of GetDBTableList

        /// <summary>
        /// Selecting the database in MySQL server. (multiple DB in one server)
        /// After the DB selected, pass the db name into GetDBTableList to get a list of table
        /// and populate them into list box     
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboDatabase_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Trace.WriteLine("ucMRVP.cs - cboDatabase_SelectedIndexChanged()");

            GetDBTableList( cboDatabase.SelectedItem.ToString() );
            btnExpID.Enabled = false; // disable button
        }//end of cboDatabase_SelectedIndexChanged

        /// <summary>
        /// Selecting different table in the list box, and reset the datagrid and context menu.
        /// 1) Binding the data table variable with data grid.
        /// 2) Creating the context menu.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstTable_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Trace.WriteLine("ucMRVP.cs - lstTable_SelectedIndexChanged");
            try
            {
                m_dtDataTable = new DataTable();			
                m_mySqlAdapter = new MySqlDataAdapter("SELECT * FROM " + lstTable.SelectedItem.ToString(), m_mySqlConn );
                m_mySqlCmdBuilder = new MySqlCommandBuilder( m_mySqlAdapter );

                m_mySqlAdapter.Fill( m_dtDataTable );
                dataGrid1.DataSource = m_dtDataTable;
          
                // dynamic building context menu
                int colCount = m_dtDataTable.Columns.Count;
                dataGrid1.CaptionText = lstTable.SelectedItem.ToString() + ": total column=" + colCount;

                dgContextMenu.MenuItems.Clear(); // clear all items in context menu
                dgContextMenu.MenuItems.Add( "Export Table", 
                    new System.EventHandler(this.MnuExportTable) ); // create even handler

                dgContextMenu.MenuItems.Add( "Export Column" );
                for( int i = 0; i < colCount; i++ )
                {
                    // Hard code index MenuItems[1], which is 2nd item of menu context (Column).
                    // The first one is export whole table to CVS. No matter how many column in submenu.
                    // Only activate the same event - export column
                    dgContextMenu.MenuItems[1].MenuItems.Add( m_dtDataTable.Columns[i].ToString(),
                        new System.EventHandler(this.subMnuExportCol)); // create event handler
                }//end of for

                SizeColumns(dataGrid1);
                EnableExportIDButton();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
        }//end of lstTable_SelectedIndexChanged

        /// <summary>
        /// Export ID Button only work for table with column name of "ServerName" and "SequenceID".
        /// Therefore check table column and enable it. Otherwise, keep disable it.
        /// </summary>
        /// <returns>true - enable; false - disable</returns>
        private bool EnableExportIDButton()
        { 
            bool rv = false;
            int idxServerName = m_dtDataTable.Columns.IndexOf( "ServerName" ); // -1: doesn't exit
            int idxSequenceID = m_dtDataTable.Columns.IndexOf( "SequenceID" ); // -1: doesn't exit
            if( (idxServerName != -1) && (idxSequenceID != -1) )
                btnExpID.Enabled = rv = true;
            else
                btnExpID.Enabled = rv = false;

            return( rv );
        }//end of EnableExportIDButton

        /// <summary>
        /// Event handler for exporting whole table to CSV file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void MnuExportTable(System.Object sender, System.EventArgs e)
        {
            Debug.WriteLine("ucMRVP.cs - MnuExportTable");

            // Create the CSV file to which grid data will be exported.
            string exportFileName = GetSaveAbsPathFileName( cboDatabase.SelectedItem.ToString() );
            try
            {   
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
        /// Event handler for exporting column to CSV file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void subMnuExportCol(System.Object sender, System.EventArgs e)
        {
            Debug.WriteLine("ucMRVP.cs - subMnuExportCol");
            System.Windows.Forms.MenuItem subMenu = (System.Windows.Forms.MenuItem)sender;

            string exportFileName = GetSaveAbsPathFileName( subMenu.Text );
            string filterExpression = "true"; // export whole column
            string sortExpression = subMenu.Text;
            ExportTableColumn( exportFileName, filterExpression, sortExpression );

        }//end of subMnuExportCol
        
        /// <summary>
        /// Get the absolute path and export file name. Default file name is the table name .csv
        /// eg: table name is mytable, file name is mytable.csv
        /// </summary>
        /// <param name="fn">input default file name</param>
        /// <returns>string: full path filename</returns>
        public string GetSaveAbsPathFileName( string fn )
        {
            string filename = "";
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
        /// Auto Size the column of datagrid
        /// </summary>
        /// <param name="grid"></param>
        protected void SizeColumns(DataGrid grid)
        {
            Graphics g = CreateGraphics();  

            DataTable dataTable = (DataTable)grid.DataSource;

            bool isContained = grid.TableStyles.Contains( dataTable.TableName );
            if( isContained )
                return;

            DataGridTableStyle dataGridTableStyle = new DataGridTableStyle();
            dataGridTableStyle.MappingName = dataTable.TableName;

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
            grid.TableStyles.Add(dataGridTableStyle);          

            g.Dispose();
        }//end of SizeColumns

        /// <summary>
        /// Export Sequence ID per server name.
        /// This is custom function that only work for table that contain "Servername", and "sequenceID" column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExpID_Click(object sender, System.EventArgs e)
        {
            DataSet ds = new DataSet();
            QATool.DataSetHelper dsHelper = new DataSetHelper( ref ds );

            try
            {
                DataTable dtDistinct = dsHelper.SelectDistinct("DistinctServerName", m_dtDataTable, "Servername" );
                int colCount = dtDistinct.Columns.Count;
                int distCount = dtDistinct.Rows.Count;

                // Get the distinct server name
                m_strArray = new string[distCount];
                for( int i = 0; i < distCount; i++ )
                {
                    m_strArray[i] = dtDistinct.Rows[i][0].ToString();
                }//end of for

                // Export the sequence ID per server name (SequenceID and ServerName hard code table column)
                for( int j = 0; j < distCount; j++ )
                {
                    string exportFileName = m_strArray[j] + ".csv";
                    string filterExpression = "ServerName='" + m_strArray[j] + "'";
                    string sortExpression = "SequenceID";

                    ExportTableColumn( exportFileName, filterExpression, sortExpression );
                }//end of for

            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
        }//end of btnExpID_Click

        /// <summary>
        /// Export a table column in acending order to a file.
        /// Filter expression: true is whole column.
        /// </summary>
        /// <param name="fileName">Full path file name</param>
        /// <param name="filterExp">The criteria to use to filter the rows</param>
        /// <param name="sortExp">A string specifying the column and sort direction</param>
        protected void ExportTableColumn( string fileName, string filterExp, string sortExp )
        {
            try
            {
                DataRow[] dataRowArray;
                // sort the data table directly
                dataRowArray = m_dtDataTable.Select(filterExp, sortExp, DataViewRowState.CurrentRows);               

                StreamWriter sw = new StreamWriter( fileName );
                int colIndex = m_dtDataTable.Columns.IndexOf(sortExp); // index of export column
                foreach( DataRow dataRow in dataRowArray )
                {
                    // Debug.WriteLine("\t " + dataRow[colIndex].ToString() );
                    sw.WriteLine( dataRow[colIndex].ToString() );
                }// end of foreach

                sw.Close();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
        }//end of ExportTableColumn

        /// <summary>
        /// Auto process milter log files to generate gap report.
        /// 1) Locate the milter log files folder (working folder)
        /// *** D O    N O T    C H E C K    I N P U T    F I L E    F O R M A T...
        /// *** M A K E    S U R E    P O I N T    T O    R I G H T    F O L D E R
        /// 2) Read files
        /// 3) Input the files into Data Table
        /// 4) Rename each log file after input into Data Table
        /// 5) Generate the gap report - Hard code d:\WUTemp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenGap_Click(object sender, System.EventArgs e)
        {
            Debug.WriteLine( "ucMRVP.cs - btnGenGap_Click ");
            try
            {
                string strFolder = ""; // log files location
                FolderBrowserDialog fbDlg = new FolderBrowserDialog();

                fbDlg.RootFolder = Environment.SpecialFolder.MyComputer; // set the default root folder
                if( fbDlg.ShowDialog() == DialogResult.OK )
                {
                    strFolder = fbDlg.SelectedPath;
                }//end of if
                else
                    return;
            
                DirectoryInfo di = new DirectoryInfo( strFolder );
	    		FileInfo[] lstFiles = di.GetFiles();
		    	int numFile = lstFiles.Length;
                if( numFile == 0 )
                {
                    MessageBox.Show( "No file in this working folder. Please select the right folder", "Warning" );
                    return;
                }//end of if

                ArrayList al = new ArrayList();
                al.Add( "X-Zantaz_Sequence:ServerName;sequenceID;DateTime" ); // add header
                for( int i = 0; i < numFile; i++ )
                {
                    StreamReader sr = File.OpenText( strFolder + "\\" + lstFiles[i].ToString() );
                    string str = sr.ReadLine(); // read the first line
                    while( str != null )
                    {
                        al.Add(str);
                        str = sr.ReadLine(); // read next line
                    }
                    sr.Close();
                }//end of for

                // allocate the array size by counting the size of arraylist
                string[] strArray = new string[al.Count];
                al.CopyTo( strArray );

                DataTable gapDataTable = ArrayToDataTable( strArray, new char[]{':',';'}, true ); // with header
                dataGrid1.DataSource = gapDataTable; 
    
                // Handle Generate the gap report per server name:
                // 1) Find out the distinct server name
                // 2) Find the gap per server along with generating the gap report
                DataSet ds = new DataSet();
                QATool.DataSetHelper dsHelper = new DataSetHelper( ref ds );

                DataTable dtDistinct = dsHelper.SelectDistinct("DistinctServerName", gapDataTable, "ServerName" );
                int colCount = dtDistinct.Columns.Count;
                int distCount = dtDistinct.Rows.Count;

                // Get the distinct server name
                if( m_strArray != null )
                    m_strArray = null; // make sure memory release
                m_strArray = new string[distCount]; // does garbage collector kick in?
                for( int i = 0; i < distCount; i++ )
                {
                    m_strArray[i] = dtDistinct.Rows[i][0].ToString();
                }//end of for

                // Export the sequence ID per server name (SequenceID and ServerName hard code table column)
                for( int j = 0; j < distCount; j++ )
                {
                    Debug.WriteLine("\t " + m_strArray[j] );
                    string exportFileName = "d:\\WUTemp\\" + m_strArray[j] + ".csv";
                    string filterExpression = "ServerName='" + m_strArray[j] + "'";
                    string sortExpression = "SequenceID";

                    GenerateGapReport( gapDataTable, exportFileName, filterExpression, sortExpression );
                }//end of for

            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch

        }//end of btnGenGap_Click

        /// <summary>
        /// This is custom function only work with column with sequential value.
        /// This function will find the missing sequence number between two integer.
        /// e.g. Given a list of number 1, 2, 3, 5, 6, 8, 9, 10
        /// Result in output file is: 4, 7 per file name.
        /// </summary>
        /// <param name="dataTable">Gap Table with multiple server name</param>
        /// <param name="fileName"></param>
        /// <param name="filterExp"></param>
        /// <param name="sortExp"></param>
        protected void GenerateGapReport( DataTable dataTable, string fileName, string filterExp, string sortExp )
        {
            Debug.WriteLine("ucMRVP.cs - GenerateGapReport"); 
            try
            {
                DataRow[] dataRowArray;
                // sort the data table directly
                dataRowArray = dataTable.Select(filterExp, sortExp, DataViewRowState.CurrentRows);               

                StreamWriter sw = new StreamWriter( fileName );
                int colIndex = dataTable.Columns.IndexOf(sortExp); // index of export column
                int iTmp = 0;
                int iNxt = 0;
                int nArray = dataRowArray.Length;
                Debug.WriteLine("\t nArray = " + nArray );

                for( int i = 0; i < nArray; i++ )
                {
                    iTmp = int.Parse(dataRowArray[i][colIndex].ToString());
                    if( (i+1) < nArray )
                        iNxt = int.Parse(dataRowArray[i+1][colIndex].ToString());
                    else 
                        break; // last element
                    Debug.WriteLine( "\t iTmp = " + iTmp + "; iNxt = " + iNxt );
                    if( (iTmp + 1) == iNxt )
                    {                        
                        iTmp = iNxt;
                    }//end of if
                    else
                    {
                        while( (iTmp+1) != iNxt )
                        {
                            Debug.WriteLine( "\t Gap:" + (iTmp+1) );                            
                            iTmp++;
                            sw.WriteLine( iTmp.ToString() ); // write to file
                        }//end of while
                    }//end of else
                }//end of for

                sw.Close();
            }//end of try
            catch( Exception ex )
            {
                Debug.WriteLine( ex.Message + "\n" + ex.GetType().ToString() + ex.StackTrace );
            }//end of catch
        }// end of GenerateGapReport

        /// <summary>
        /// Converts an array of strings to a DataTable
        /// This method can be used with any type of single character delimiter 
        /// and multiple single character delimiters, but does not ignore delimiters 
        /// if they are inside quotation marks. e.g. Input
        /// string[] list = {"Symbols","IBM","INTC","DELL","SUNW","MSFT"};
        /// DataTable dt = ArrayToDataTable(list, new char[1]{','}, true);
        /// Output: a column of "Symbols","IBM","INTC","DELL","SUNW","MSFT"
        /// </summary>
        /// <param name="strArray"></param>
        /// <param name="delimiter">e.g: new char[]{';',':'} </param>
        /// <param name="bHeaders">Indicate does the first line of file contain column header</param>
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
                if( bHeaders ) // the first line of the file
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


	}
}
