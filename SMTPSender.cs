using System;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;

namespace QATool
{
    /// <summary>
    /// Custom exception class for SMTPSender
    /// </summary>
    public class SMTPException:System.Exception
    {
        private string _msg = "";

        public SMTPException( string str )
        {
            _msg = str;
        }

        /// <summary>
        /// Exception Message
        /// </summary>
        public string SmtpMessage
        {
            get
            {
                return _msg;
            }
            set
            {
                _msg = value;
            }
        }// end of construction
    }// end of class - SMTPException

	/// <summary>
	/// Class SMTPSender inherites from System.net.Socket.TcpClient that provides all the basic
	/// functionality to do TCP/IP programming.
	/// </summary>
	public class SMTPSender:System.Net.Sockets.TcpClient
	{
        private string _strFrom    = "";
        private string _strTo      = "";
        private string _strCC      = "";
        private string _strBCC     = "";
        private string _strSubject = "";
        private string _strBodyTxt = "";
        private string _strSentDay = ""; 
        private string _strServer  = ""; // SMTP server name or ip
        private string _strPortNum = "";
        private string _strContent = "";

        private const int BYTE_SIZE = 8192;

		public SMTPSender()
		{
			
		}// end of constructor

        /// <summary>
        /// RFC 821 - MAIL FROM:
        /// </summary>
        public string mailFrom
        {
            set
            {
                _strFrom = value;
            }
        }//end of property - mailFrom

        /// <summary>
        /// RFC 821 - RCPT TO:
        /// May contain more than one mail address with semicolon or comma. 
        /// </summary>
        public string mailTo
        {
            set
            {
                _strTo = value;
            }
        }// end of property

        /// <summary>
        /// RFC 822 - CC
        /// </summary>
        public string mailCC
        {
            set
            {
                _strCC = value;
            }
        }// end of property

        /// <summary>
        /// RFC 822 - BCC
        /// </summary>
        public string mailBCC
        {
            set
            {
                _strBCC = value;
            }
        }//end of property

        /// <summary>
        /// RFC 822 - Mail Send date
        /// </summary>
        public string mailSentDate
        {
            set
            {
                _strSentDay = value;
            }
        }//end of property

        /// <summary>
        /// RFC 822 - subject line - take whatever from user.
        /// </summary>
        public string mailSubject
        {
            set
            {
                _strSubject = value;
            }
        }//end of property

        /// <summary>
        /// RFC 822 - mail body
        /// </summary>
        public string mailBody
        {
            set
            {
                _strBodyTxt = value;
            }
        }//end of property

        /// <summary>
        /// SMTP Server name or IP
        /// </summary>
        public string smtpServer
        {
            get
            {
                return _strServer;
            }
            set
            {
                _strServer = value;
            }
        }//end of property

        /// <summary>
        /// SMTP Server Port number
        /// </summary>
        public string smtpPortNum
        {
            get
            {
                return _strPortNum;
            }
            set
            {
                // TO DO : validate input - numeric characters only - regular expression
                _strPortNum = value;
            }
        }//end of property

        /// <summary>
        /// RFC 822 - Set Content type - for unparsable content
        /// </summary>
        public string mailContentType
        {
            set
            {
                _strContent = value;
            }
        }//end of property - mailContentType

        /// <summary>
        /// Does all the work: Initiates SMTP communication, send the mail.
        /// This method uses three methods: 
        /// Connect() - inherited from the TcpClient class for establishing a TCP connection between client and TCP server.
        /// WriteToSocket() - write data to socket in ASCII format.
        /// ReadFromSocket() - read socket stream using GetStream method in TcpClient class
        /// Mail building based on property value... (similar to MS API)
        /// </summary>
        public void SmtpSend()
        {
            Trace.WriteLine("SMTPSender.cs - SmtpSend()" );
            string strMsg;
            string strReply; // reply from smtp server

            try
            {
                Debug.WriteLine("\tSmtpSend() - Connect to smtp server");
                Connect( _strServer, int.Parse(_strPortNum) ); // inherit from System.Net.Sockets.TcpClient
                strReply = ReadFromSocket();
                if( strReply.Substring(0,3) != "220" )
                    throw new SMTPException( strReply ); // will catch in the caller (eg. main)

                Debug.WriteLine("\tSmtpSend() - Test the connection HELO");
                strMsg = "HELO world\r\n"; // test connection
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );

                Debug.WriteLine("\tSmtpSend() - write mail from into socket");
                strMsg = "MAIL FROM: " + _strFrom + "\r\n"; // set up 821 header
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0,3) != "250" )
                    throw new SMTPException( strReply );

                Debug.WriteLine("\tSmtpSend() - write rcpt to into socket");
                strMsg = "RCPT TO: " + _strTo + "\r\n";
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );

                if( _strCC != "" ) // CC exist - show in different repository but only one physical mail
                {
                    Debug.WriteLine("\tSmtpSend() - write CC (rcpt to) into socket");
                    strMsg = "RCPT TO: " + _strCC + "\r\n";
                    WriteToSocket( strMsg );
                    strReply = ReadFromSocket();
                    if( strReply.Substring(0, 3) != "250" )
                        throw new SMTPException( strReply );
                }//end of if - CC exist

                if( _strBCC != "" ) // BCC exist - show in different repository but only one physical mail
                {
                    Debug.WriteLine("\tSmtpSend() - write BCC (rcpt to) into socket");
                    strMsg = "RCPT TO: " + _strBCC + "\r\n";
                    WriteToSocket( strMsg );
                    strReply = ReadFromSocket();
                    if( strReply.Substring(0, 3) != "250" )
                        throw new SMTPException( strReply );
                }//end of if - BCC exist

                Debug.WriteLine("\tSmtpSend() - Write DATA into socket - signaling SMTP server");
                strMsg = "DATA\r\n";
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "354" )
                    throw new SMTPException( strReply );

                strMsg = "Subject: " + _strSubject + "\r\n"; // start 822 header set up
                strMsg += "To: "     + "Identical" + "\r\n"; // Everything here is identical
                strMsg += "CC: "     + "Faking CC" + "\r\n"; // BCC is not require..
                strMsg += "From: "   + _strFrom    + "\r\n";
                strMsg += "Date: "   + _strSentDay + "\r\n";
                strMsg += "MIME-Version: 1.0\r\n";
                strMsg += "Content-Type: " + _strContent + "\r\n"
//                    + "Content-Type: text/html;\r\n"
                    + "charset=\"iso-8859-1\"\r\n"
                    + "\r\n"; // blink line
                strMsg += _strBodyTxt + "\r\n";
                strMsg += ".\r\n"; // period - end of mail

                Debug.WriteLine("\tSmtpSend() - write message (DATA) into socket");
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );
                Debug.WriteLine("\tSmtpSend() - Send now and quit");
                strMsg = "QUIT\r\n"; // Send now...
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.IndexOf("221") == -1 )
                    throw new SMTPException( strReply );

                Close(); // TCP connection - inherited from System.Net.Sockets.TcpClient
            }//end of try
            catch( SocketException ex )
            {
                Trace.WriteLine("\tSocket Exception: " + ex.Message.ToString() );
            }//end of catch - socket exception
        }//end of smtpSend


        /// <summary>
        /// Does all the work: Initiates SMTP communication, send the mail.
        /// This method uses three methods: 
        /// Connect() - inherited from the TcpClient class for establishing a TCP connection between client and TCP server.
        /// WriteToSocket() - write data to socket in ASCII format.
        /// ReadFromSocket() - read socket stream using GetStream method in TcpClient class
        /// Generic streaming a file into socket... Use to send RFC822 file is perfect or eml...
        /// </summary>
        public void SmtpSend( string fileName )
        {
            Trace.WriteLine("SMTPSender.cs - SmtpSend( fileName )" );
            string strMsg;
            string strReply; // reply from smtp server

            StreamReader sr = null;
            try
            {
                Debug.WriteLine("\tSmtpSend() - Connect to smtp server");
                Connect( _strServer, int.Parse(_strPortNum) ); // inherit from System.Net.Sockets.TcpClient
                strReply = ReadFromSocket();
                if( strReply.Substring(0,3) != "220" )
                    throw new SMTPException( strReply ); // will catch in the caller (eg. main)

                Debug.WriteLine("\tSmtpSend() - Test the connection HELO");
                strMsg = "HELO world\r\n"; // test connection
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );

                Debug.WriteLine("\tSmtpSend() - write mail from into socket");
                strMsg = "MAIL FROM: " + _strFrom + "\r\n"; // set up 821 header
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0,3) != "250" )
                    throw new SMTPException( strReply );

                Debug.WriteLine("\tSmtpSend() - write rcpt to into socket");
                strMsg = "RCPT TO: " + _strTo + "\r\n";
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );

                Debug.WriteLine("\tSmtpSend() - Write DATA into socket - signaling SMTP server");
                strMsg = "DATA\r\n";
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "354" )
                    throw new SMTPException( strReply );

                sr = new StreamReader( fileName );
                while( (strMsg = sr.ReadLine()) != null )
                {
                    Debug.WriteLine( "\tLooping the text file line by line");

                    // special case for Message-ID
                    if( strMsg.IndexOf( "<KENTEST>" ) != -1 )
                    {
                        strMsg = "Message-ID: <" + System.Guid.NewGuid().ToString() + "SMTPSender.cs>";
                    }

                    WriteToSocket( strMsg + "\r\n" );
                    // I think don't need the new line here
//                    WriteToSocket( strMsg );
                }//end of while

                strMsg = ".\r\n"; // period - end of mail
                // since no new line in message body add here
//                strMsg = "\r\n.\r\n"; // period - end of mail
                Debug.WriteLine("\tSmtpSend() - end of mail: write dot into socket");
                WriteToSocket( strMsg );

                strReply = ReadFromSocket();
                if( strReply.Substring(0, 3) != "250" )
                    throw new SMTPException( strReply );
                Debug.WriteLine("\tSmtpSend() - Send now and quit");
                strMsg = "QUIT\r\n"; // Send now...
                WriteToSocket( strMsg );
                strReply = ReadFromSocket();
                if( strReply.IndexOf("221") == -1 )
                    throw new SMTPException( strReply );

                Close(); // TCP connection - inherited from System.Net.Sockets.TcpClient
            }//end of try
            catch( SocketException ex )
            {
                Trace.WriteLine("\tSocket Exception: " + ex.Message.ToString() );

            }//end of catch - socket exception
            catch( IOException ioex )
            {
                Trace.WriteLine("\tIO Exception: " + ioex.Message.ToString() );
            }// end of catch - IO exception
            catch( Exception gEx )
            {
                Trace.WriteLine("\tGeneric Exception: " + gEx.Message.ToString() );
            }// end of catch - IO exception

        }//end of smtpSend

        /// <summary>
        /// Write data to socket in ASCII format. dotNet string class is unicode. ie: need to convert to ASCII encoding.
        /// </summary>
        public void WriteToSocket( string msg )
        {
            Trace.WriteLine("SMTPSender.cs - WriteToSocket():" + msg.ToString() );
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            byte[] writeBuffer = new byte[BYTE_SIZE]; // 8K
            writeBuffer = asciiEncoding.GetBytes( msg );

            NetworkStream nwStream = GetStream();
            nwStream.Write( writeBuffer, 0, writeBuffer.Length );

        }//end of WriteToSocket

        /// <summary>
        /// Read data stream from socket and convert ASCII back to native dotNet string.
        /// </summary>
        /// <returns></returns>
        public string ReadFromSocket()
        {
            Trace.WriteLine( "SMTPSender.cs - ReadFromSocket()" );
            System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
            byte[] serverBuffer = new byte[BYTE_SIZE]; // 8K

            NetworkStream nwStream = GetStream();
            int count = nwStream.Read( serverBuffer, 0, serverBuffer.Length );
            if( count == 0 ) // no more data
                return( "" );
            else
                return( asciiEncoding.GetString(serverBuffer, 0, count) );
        }//end of ReadFromSocket
	}//end of class - SMTPSender
}// end of namespace - QATool
