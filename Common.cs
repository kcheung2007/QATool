using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Web;
using System.Web.Mail;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for Common.
	/// No and CANNOT have local variable. If object exist only for implementation purpose.
	/// NO data should store here. ONLY method for accessing common function.
	/// </summary>
	public class CommObj
	{
        private string _logPath = Application.StartupPath;

		public CommObj()
		{
			//
			// TODO: Add constructor logic here
			//
		}


		/// <summary>
		/// Read addresses from a text file
		/// 1) Get file name from open file dialog
		/// 2) Read the file line by line
		/// 3) Append ; & space to form one line
		/// 4) Put the result back to input parameter		
		/// </summary>
		public void LoadAddrFromFile(ref string inStr)
		{
			// Show the Open File Dialog.
			// If the user clicked OK in the dialog and a text file was selected, open it.
			// Display an OpenFileDialog and show the read-only files...
			OpenFileDialog ofDlg = new OpenFileDialog();
			ofDlg.ShowReadOnly = true;

			if( ofDlg.ShowDialog() == DialogResult.OK )
			{
				if( ofDlg.FileName != "" )
				{
					inStr = ""; // clear the default text
					StreamReader sr = new StreamReader( ofDlg.FileName );
					String str;
					while( (str = sr.ReadLine()) != null )
					{
						inStr = inStr + str + "; ";
					}//end of while
					sr.Close();
				}//end of if - open file name
			}// end of if - open file dialog				
		}//end of LoadAddrFromFile

		/// <summary>
		/// Open a file and add a line of text into a text file.
		/// If file doesn't exist, create it. Will log the time stamp
		/// </summary>
		/// <param name="fileName"></param>
		/// <param name="strText">text to add into file</param>
//		public void LogToFile( string fileName, string strText, bool bdebug )
        public void LogToFile( string fileName, string strText )
		{
//            if( !bdebug )
//                return;
			/*** save for reference
			 * string logFile = DateTime.Now.ToShortDateString().Replace(@"/",@"-").Replace(@"\",@"-" + ".log";
			 * save for reference */
			
			StreamWriter sw = File.AppendText( fileName);
			Log(strText, sw);

	        // Close the writer and underlying file.
		    sw.Close();
	    }// end of LogToFile

		/// <summary>
		/// Overload function. Will log the time stamp
		/// </summary>
		/// <param name="strText">text to add into file</param>
//		public void LogToFile( string strText, bool bdebug )
        public void LogToFile( string strText )
		{
//            if( !bdebug )
//                return;

			LogToFile( _logPath + "\\QATool.log", strText );
		}// end of LogToFile
		
		public static void Log( String logMessage, TextWriter w )
		{
			w.Write("Log: "); // if want blank line, add /r/n -> w.Write("/r/nLog Entry : ");
			w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToShortDateString());
			w.WriteLine("{0}", logMessage);
//			w.WriteLine ("-------------------------------");
			// Update the underlying file.
			w.Flush(); 
		}// end of Log

		/// <summary>
		/// Special function for logging GUID ONLY!!!
		/// </summary>
		/// <param name="fileName"></param>
		/// <param name="strGUID"></param>
		public void LogGUID( string fileName, string strGUID )
		{
			StreamWriter sw = File.AppendText( fileName);
			sw.WriteLine("{0}",strGUID );
			sw.Close();
		}//end of LogGUID

		/// <summary>
		/// Write to a file line by line. No time stamp
		/// </summary>
		/// <param name="fn">file name</param>
		/// <param name="line">text line</param>
		public void WriteLineByLine(string fn, string line)
		{
			StreamWriter sw = File.AppendText( fn );
			sw.WriteLine("{0}", line );
			sw.Close();		
		}// end of WriteLineByLine

		/// <summary>
		/// Test the smtp server connection
		/// </summary>
		/// <param name="smtpHost">SMTP server name or IP address</param>
		/// <param name="smtpPort">SMTP server listening port number - default 25</param>
		/// <returns>true - connection OK</returns>
		/// <returns>false - connection FAIL</returns>
		public bool TestSMTPConnection(string smtpHost, string smtpPort)
		{
			bool rv = false; // no connection
			System.Net.Sockets.TcpClient tcpClient = new TcpClient();
			try
			{	// If SmtpServer is not set, local SMTP server is used
				SmtpMail.SmtpServer = smtpHost;

				// Test the connection - if smtp server down, exit.				
				tcpClient.Connect( SmtpMail.SmtpServer, int.Parse(smtpPort) );
				
				rv = true; // connection OK, no exception							
			}//end of try
			catch( SocketException ex )
			{
				rv = false; // connection FAIL
				MessageBox.Show(ex.Message.ToString(), "SMTP Connection Test");
			}//end of catch - SocketException
			finally
			{				
				if( tcpClient != null )
				{
					Debug.WriteLine("Finally - close TCP Client connection");
					tcpClient.Close();
				}// end of if
			}//end of finally
			return( rv );
		}//end of TestSMTPConnection

		/// <summary>
		/// Test the sql server DB connection
		/// 1) Open connection. is there any exception?
		/// 2) Close the connection if and only if no exception
		/// </summary>
		/// <param name="strConnetion">fully constructed connection string</param>
		/// <returns>true - connection OK</returns>
		/// <returns>false - connection FAIL</returns>
		public bool TestSQLConnection(string strConnect)
		{
			bool rv = false; // no connection
			try
			{
				SqlConnection myConnection = new SqlConnection(strConnect);
				myConnection.Open();
				MessageBox.Show("ServerVersion: " + myConnection.ServerVersion
					+ "\nState: " + myConnection.State.ToString()
					+ "\nData Base " + myConnection.Database.ToString()
					+ "\nData Source " + myConnection.DataSource.ToString()
					+ "\n Connection OK", "DB Connection Test");
				myConnection.Close();
		
				rv = true; // every OK
			}//end of try
			catch( SqlException sqlEx )
			{
				string errorMessages = "";
				for (int i=0; i < sqlEx.Errors.Count; i++)
				{
					errorMessages += "Index #" + i + "\n" +
						"Message: " + sqlEx.Errors[i].Message + "\n" +
						"LineNumber: " + sqlEx.Errors[i].LineNumber + "\n" +
						"Source: " + sqlEx.Errors[i].Source + "\n" +
						"Procedure: " + sqlEx.Errors[i].Procedure + "\n";
				}
				MessageBox.Show( errorMessages, "SQL Error" );
				
				rv = false; // something wrong
			}//end of catch 0 SQLException

			return( rv );
		}// emd pf TestSQLConnection

		/// <summary>
		/// Read from file line by line and Load the item into ComboBox.
		/// </summary>
		/// <param name="cbx">Pass by reference - combobox control</param>
		/// <param name="fn">File Name that contains combo box item</param>
		public void LoadComboBoxItem( ComboBox cbx, string fn )
		{
			String str;
			StreamReader sr = new StreamReader( fn );
			while( (str = sr.ReadLine()) != null )
			{
				cbx.Items.Add( str );
			}//end of while
			sr.Close();
		}//end of LoadComboBoxItem

		/// <summary>
		/// Initial the ComboBoxItem during start up.
		/// Every call will loop through the whole ini file QATool.ini for finding 
		/// the specific section.
		/// 1) Ensure the QATool.ini exit.
		/// 2) Find the section.
		/// 3) Add item for that particular section until '[' find.
		/// 4) Ignor line start with # sign
		/// </summary>
		/// <param name="cbx"></param>
		/// <param name="iniSection"></param>
		/// <param name="iniFileName"></param>
        public void InitComboBoxItem(ComboBox cbx, string iniSection, string iniFileName)
        {
            string iniFile = iniFileName;
            string line = "";
            bool   done = false;
            try
            {
                cbx.Items.Clear();
                StreamReader sr = new StreamReader( iniFile );
                while( (line = sr.ReadLine()) != null && !done )
                {
                    line = line.Trim(' ');
                    if( (line != "") && (line[0] != '#') )
                    {
                        if( line.IndexOf(iniSection) != -1 )
                        {
                            while( (line = sr.ReadLine()) != null ) // read next line
                            {
                                line = line.Trim(' ');
                                if( (line != "") && line[0] != '[' ) // DONE if empty line or  [ 
                                    cbx.Items.Add( line );
                                else
                                {
                                    done = true;
                                    break; // inner while
                                }
                            }//end of while
                        }//end of if
                    }//end of if - ignor #						
                }//end of while 
                sr.Close();
            }//end of try 
            catch( IOException ioEx )
            {
                Debug.WriteLine( "\t" + ioEx.Message );
            }//end of catch IOException

        }//end of InitComboBoxItem

        /// <summary>
        /// Initial the ComboBoxItem during start up.
        /// Every call will loop through the whole ini file QATool.ini for finding 
        /// the specific section.
        /// 1) Ensure the QATool.ini exit.
        /// 2) Find the section.
        /// 3) Add item for that particular section until '[' find.
        /// 4) Ignor line start with # sign
        /// </summary>
        /// <param name="cbx"></param>
        /// <param name="iniSection"></param>
        public void InitComboBoxItem(ComboBox cbx, string iniSection)
        {
            Trace.WriteLine("Common.cs - InitComboBoxItem( ComboBox, string )");
            InitComboBoxItem( cbx, iniSection, getFullPathIni() );
        }//end of InitComboBoxItem

		/// <summary>
		/// Get the full path file name of QATool.ini...
		/// 1) Check the current path first.
		/// 2) Not exit, pop up a dialog for location.
		/// </summary>
		/// <returns></returns>
		public string getFullPathIni()
		{
			Trace.WriteLine("Common.cs - getFullPathIni");
			string path = Directory.GetCurrentDirectory();
			string uri = path + "\\QATool.ini"; // uri = Full Path + file name + ext

			if( !File.Exists(uri) )
			{
				OpenFileDialog ofDlg = new OpenFileDialog();
				ofDlg.ShowReadOnly = true;
//				ofDlg.RestoreDirectory = true;

				if( ofDlg.ShowDialog() == DialogResult.OK )
				{
					if( ofDlg.FileName != "" )
					{
						uri = ofDlg.FileName;
					}//end of if - open file name
				}// end of if - open file dialog								
			}// end of if - File !exist
			
			return( uri );;
		}// end of getFullPathIni

        /// <summary>
        /// Mainly check does file exist, and get file info
        /// </summary>
        /// <param name="fn"></param>
        /// <returns>0 - normal, 1 - file doesn't exist</returns>
        public int CheckFileInfo( string fn )
        {
            Trace.WriteLine("Common.cs - CheckFileInfo: " + fn);

            int rv = 0; // assume normal
            string str;
            if( fn != "" )
            {
                try
                {
                    FileInfo fInfo  = new FileInfo( fn );

                    str = "File Name = " + fInfo.FullName + "\n"
                        + "Creation Time = " + fInfo.CreationTime + "\n"
                        + "Last Access Time = " + fInfo.LastAccessTime + "\n"
                        + "Last Write Time = " + fInfo.LastWriteTime + "\n"
                        + "Size = " + fInfo.Length;
                }// end of try
                catch( Exception ex )
                {
                    Trace.WriteLine( ex.Message.ToString() );
                    str = ex.Message.ToString();
                    rv = 1;
                }// end of catch
            }
            else
            {
                str = "No File Name specified!!";
                rv = 1;
            }// end of else
            return( rv );
        }// end of CheckFileInfo

        /// <summary>
        /// Full Qaulified Path for the log file. Default is the application start up path
        /// </summary>
        public string logFullPath
        {
            get
            {
                return( _logPath );
            }
            set
            {
                _logPath = value;
            }
        }
	}// end of class - CommObj
}
