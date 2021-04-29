using System;
using System.Diagnostics;

namespace QATool
{

	/// <summary>
	/// Store all the URL input parameters
	/// Get the real copy value. (NOT by reference)
	/// 
	/// example of Query Search URL: 
	/// https://zantaz.digitalsafe.net/zantaz/DigitalSafe/QueryTypeSelectorServlet
	/// ?SessionKey=2WHq0y95Mo_11qwVyP9jf_RTbZbGbrVLUFAflTJ_Xy8%3D
	/// &NewSearch=new
	/// &DocumentClassList=D0000001
	/// &repositoryID=qES93i5AUCuiI46BA8VjeQ%3D%3D
	/// &documentClassList=D0000001
	/// &UnsupportedContentSearch=disabled
	/// &D_FROM=tin*
	/// &FROM=tin*
	/// &D_APPARENTLY-TO=tin*
	/// &APPARENTLY-TO=
	/// &D_SUBJECT=
	/// &SUBJECT=
	/// &D_%3DBODY%3D=
	/// &%3DBODY%3D=
	/// &D_%3DATTACHMENTNAME%3D=
	/// &%3DATTACHMENTNAME%3D=
	/// &D_MESSAGE-ID=
	/// &MESSAGE-ID=
	/// &D_X-ZANTAZ=
	/// &X-ZANTAZ=
	/// &amo=8
	/// &ady=22
	/// &ayr=2003
	/// &ahour=0
	/// &amin=0
	/// &bmo=8
	/// &bdy=30
	/// &byr=2003
	/// &bhour=0
	/// &bmin=0
	/// &tz=GMT
	/// 
	/// example of Message Display URL:
	/// http://10.1.89.201/zantaz/DigitalSafe/MsgDisplayServlet
	/// ?SessionKey=_drGeLtnW9wu13JJqB8gqPAfso0LGUy3
	/// &Navigating=true
	/// &QueryStatus=N/A
	/// &domain=company1.zantaz.com
	/// &smgID=jbY0SDLElx-XXouwmdWKppdei7CZ1Yqm1r2kRfHVfdpt65G9srzAAdiCk-w8TdRWHIR2szexeh8FAv0s71wWLA==
	/// &msgPath=UpjNXei_1669IAHq4xxrhZdei7CZ1YqmgJ-LJWF-xQBqc-dXS8j3TuLlQUTMaace0IZ7QQ2HXoCXXouwmdWKpna324C-5qhDsB9eIIscppvQKxa34VKDN2ZU6AZT0KwCvvupainiKVkM74l7sTo_Nw==
	/// &offset=0
	/// &ThumbIndex=0
	/// &ShowDetails=false
	/// &UnparseableDefaultDate=1063038579000
	/// &isUnIndexableDocument=false
	/// &selMsgs=
	/// &all=0
	/// </summary>
	public class URLDataObj
	{
		// define const in case the key name change
		private char[] token = {'&', '>', '"'};

		// For query search url
		public const string EQUAL_SIGN = "%3D";
		public const string SESSION_KEY = "SessionKey=";
		public const string NEW_SEARCH = "NewSearch=";
		public const string DOC_CLASS_LIST = "DocumentClassList=";
		public const string REPOSITORY_ID = "repositoryID=";
		public const string DOC_CLASS_LIST2 = "documentClassList=";
		public const string U_CONTENT_SEARCH = "UnsupportedContentSearch=";
		public const string FROM = "FROM=";
		public const string APPARENTLY_TO = "APPARENTLY-TO=";
		public const string SUBJECT = "SUBJECT=";
		public const string BODY = "%3DBODY%3D=";
		public const string ATTACH_NAME = "%3DATTACHMENTNAME%3D=";
		public const string MESSAGE_ID = "MESSAGE-ID=";
		public const string X_ZANTAZ = "X-ZANTAZ=";
		public const string AMO = "amo=";
		public const string ADY = "ady=";
		public const string AYR = "ayr=";
		public const string AHOUR = "ahour=";
		public const string AMIN = "amin=";
		public const string BMO = "bmo=";
		public const string BDY = "bdy=";
		public const string BYR = "byr=";
		public const string BHOUR = "bhour=";
		public const string BMIN = "bmin=";
		public const string TZ = "tz=";
		
		private string _BaseURL = "";
		private string _SessionKey = "";
		private string _DocumentClassList = "";
		private string _NewSearch = "new"; // re-check: is always "new"?
		private string _RepositoryID = "";
		private string _UnspportedContent = "disabled"; // re-check: is always "disabled"?
		private string _From = "";
		private string _To = "";
		private string _Subject = "";
		private string _Body = "";
		private string _AttachName = "";
		private string _MessageID = "";
        private string _XZantaz = "";
		private string _StartMonth = "";
		private string _StartDay = "";
		private string _StartYear = "";
		private string _StartHour = "";
		private string _StartMinute = "";
		private string _EndMonth = "";
		private string _EndDay = "";
		private string _EndYear = "";
		private string _EndHour = "";
		private string _EndMinute = "";
		private string _TimeZone = "GMT";

		// For message display url
		public const string NAVIGATING = "Navigation=";
		public const string QUERY_STATUS = "QueryStatus=";
		public const string DOMAIN = "domain=";
		public const string SMG_ID = "smgID=";
		public const string MSG_PATH = "msgPath=";
		public const string OFFSET = "offset=";
		public const string THUMB_INDEX = "ThumbIndex=";
		public const string SHOW_DETAILS = "ShowDetails=";
		public const string UNPARSE_DATE = "UnparseableDefaultDate=";
		public const string IS_UNIDX_DOC = "isUnIndexableDocument=";
		public const string FORMAT_DATE = "formattedDate=";
		public const string UNIDX_SUBJECT = "UnIndexableSubject=";
		public const string SEL_MSGS = "selMsgs=";
		public const string ALL = "all=";

		private string _Navigating = "";
		private string _QueryStatus = "";
		private string _Domain = "";
		private string _SmgID = "";
		private string _MsgPath = "";
		private string _Offset = "";
		private string _ThumbIndex = "";
		private string _ShowDetails = "";
		private string _UnparseDate = "";
		private string _IsUnidxDoc = "";
		private string _FormatDate = "";
		private string _UnIdxSubject = "";
		private string _SelectMsgs = "";
		private string _All = "";

		public URLDataObj()
		{
			// TODO: Add constructor logic here			
		}//end of constructor

		public string getBaseURL()
		{
			return( _BaseURL );
		}//end of getBaseURL

		public void setBaseURL( string str )
		{
			_BaseURL = str;
		}//end of setBaseURL

		public string getSessionKey()
		{
			return( _SessionKey );
		}//end of getSessionKey

		public void setSessionKey( string str )
		{
			_SessionKey = str;
		}//end of setSessionKey

		public string getNewSearch()
		{
			return( _NewSearch );
		}//end of getNewSearch

		public void setNewSearch( string str )
		{
			_NewSearch = str;
		}//end of setNewSearch

		public string getDocumentClassList()
		{
			return( _DocumentClassList );
		}//end of getDocumentClassList

		public void setDocumentClassList( string str )
		{
			_DocumentClassList = str;
		}// end of setSocumentClassList
		
		public string getRepositoryID()
		{
			return( _RepositoryID );
		}//end of getRepositoryID

		public void setRepositoryID( string str )
		{
			_RepositoryID = str;
		}//end of setRepositoryID

		public string getUnContentSearch()
		{
			return( _UnspportedContent );
		}//end of getUnContentSearch

		public void setUnContentSearch( string str )
		{
			_UnspportedContent = str;
		}//end of setUnContentSearch

		public string getFrom()
		{
			return( _From );
		}//end of getFrom

		public void setFrom( string str )
		{
			_From = str;
		}//end of setFrom

		public string getTo()
		{
			return( _To );
		}//end of getTo

		public void setTo( string str )
		{
			_To = str;
		}//end of setTo

		public string getSubject()
		{
			return( _Subject );
		}//end of getSubject

		public void setSubject( string str )
		{
			_Subject = str;
		}//end of setSubject

		public string getBody( )
		{
			return( _Body );
		}//end of getBody

		public void setBody( string str )
		{
			_Body = str;
		}//end of setBody

		public string getAttachName()
		{
			return( _AttachName );
		}//end of getAttachName

		public void setAttachName( string str )
		{
			_AttachName = str;
		}// end of setAttachName

		public string getMessageID()
		{
			return( _MessageID );
		}//end of getMessageID

		public void setMessageID( string str )
		{
			_MessageID = str;
		}//end of setMessageID

        public string getXZantaz()
        {
            return( _XZantaz );
        }//end of getXZantaz

        public void setXZantaz( string str )
        {
            _XZantaz = str;
        }//end of setXZantaz

        public string getStartMonth()
		{
			return( _StartMonth );
		}//end of getStartMonth

		public void setStartMonth( string str )
		{
			_StartMonth = str;
		}//end of setStartMonth

		public string getStartDay()
		{
			return( _StartDay );
		}//end of getStartDay

		public void setStartDay( string str )
		{
			_StartDay = str;
		}//end of setStartDay

		public string getStartYear()
		{
			return( _StartYear );
		}//end of getStartYear

		public void setStartYear( string str )
		{
			_StartYear = str;
		}//emd of setStartYear

		public string getStartHour()
		{
			return( _StartHour );
		}//end of getStartHour

		public void setStartHour( string str )
		{
			_StartHour = str;
		}// end of setStartHour

		public string getStartMinute()
		{
			return( _StartMinute );
		}//end of getStartMinute

		public void setStartMinute( string str )
		{
			_StartMinute = str;
		}//end of setStartMinute

		public string getEndMonth()
		{
			return( _EndMonth );
		}//end of getEndMonth

		public void setEndMonth( string str )
		{
			_EndMonth = str;
		}//end of setEndMonth

		public string getEndDay()
		{
			return( _EndDay );
		}//end of getEndDay

		public void setEndDay( string str )
		{
			_EndDay = str;
		}//end of setEndDay

		public string getEndYear()
		{
			return( _EndYear );
		}//end of getEndYear

		public void setEndYear( string str )
		{
			_EndYear = str;
		}//end of setEndYear

		public string getEndHour()
		{
			return( _EndHour );
		}//end of getEndHour

		public void setEndHour( string str )
		{
			_EndHour = str;
		}//end of setEndHour

		public string getEndMinute()
		{
			return( _EndMinute );
		}//end of getEndMinute

		public void setEndMinute( string str )
		{
			_EndMinute = str;
		}//end of setEndMinute

		public string getTimeZone()
		{
			return( _TimeZone );
		}//end of getTimeZone

		public void setTimeZone( string str )
		{
			_TimeZone = str;
		}//end of setTimeZone

		public string getNavigating()
		{
			return( _Navigating );
		}//end of getNavigating

		public void setNavigating( string str )
		{
			_Navigating = str;
		}//end of setNavigating

		public string getQueryStatus()
		{
			return( _QueryStatus );
		}//end of getQueryStatus

		public void setQueryStatus( string str )
		{
			_QueryStatus = str;
		}//end of setQueryStatus

		public string getDomain()
		{
			return( _Domain );
		}//end of getDomain

		public void setDomain( string str )
		{
			_Domain = str;
		}//end of setDomain

		public string getSmgID()
		{
			return( _SmgID );
		}//end of getSmgID

		public void setSmgID( string str )
		{
			_SmgID = str;
		}//end of setSmgID

		public string getMsgPath()
		{
			return( _MsgPath );
		}//end of getMsgPath

		public void setMsgPath( string str )
		{
			_MsgPath = str;
		}//end of setMsgPath

		public string getOffset()
		{
			return( _Offset );
		}//end of getOffset

		public void setOffset( string str )
		{
			_Offset = str;
		}//end of setOffset

		public string getThumbIndex()
		{
			return( _ThumbIndex );
		}//end of getThumbIndex

		public void setThumbIndex( string str )
		{
			_ThumbIndex = str;
		}//end of setThumbIndex

		public string getShowDetails()
		{
			return( _ShowDetails );
		}//emd of getShowDetails

		public void setShowDetails( string str )
		{
			_ShowDetails = str;
		}//end of setShowDetails

		public string getUnparseDate()
		{
			return( _UnparseDate );
		}//end of getUnparseDate

		public void setUnparseDate( string str )
		{
			_UnparseDate = str;
		}//end of setUnparseDate

		public string getIsUnidxDoc()
		{
			return( _IsUnidxDoc );
		}//end of getIsUnidxDoc

		public void setIsUnidxDoc( string str )
		{
			_IsUnidxDoc = str;
		}//end of setIsUnidxDoc

		public string getSelectMsgs()
		{
			return( _SelectMsgs );
		}//end of getSelectMsgs

		public void setSelectMsgs( string str )
		{
			_SelectMsgs = str;
		}//end of setSelectMsgs

		public string getAll()
		{
			return( _All );
		}//end of getAll

		public void setAll( string str )
		{
			_All = str;
		}//end of setAll

		public string getFormatDate()
		{
			return( _FormatDate );
		}//end of getFormatDate

		public void setFormatDate( string str )
		{
			_FormatDate = str;
		}//end of setFormatDate

		public string getUnIdxSubject()
		{
			return( _UnIdxSubject );
		}//end of getUnIdxSubject

		public void setUnIdxSubject( string str )
		{
			_UnIdxSubject = str;
		}//end of setUnIdxSubject

		/// <summary>
		/// Pass in a line of HTML login page and parse the SessionKey/
		/// Format is pre-defined. Othwise, will not work.
		/// </summary>
		/// <param name="str">A line in HTML file</param>
		public void ExtractSessionKey( string str )
		{
			int idx; 
			int iStart;
			int iEnd; // end of substring index
			if( (idx = str.IndexOf(SESSION_KEY)) != -1 ) // -1 not found
			{
				if( (iStart = str.IndexOf('=', idx)) != -1 )
				{
					iStart++; // move to next; next of =
					iEnd = str.IndexOfAny( token, iStart );
					_SessionKey = str.Substring( iStart, iEnd-iStart );
				}
			}//end of if
		}// end of ExtractSessionKey

		/// <summary>
		/// Pass in a line of HTML login page and parse the DocumentClassList.
		/// Format is pre-defined. Othwise, will not work.
		/// </summary>
		/// <param name="str"></param>
		public void ExtractDocumentClassList( string str )
		{
			int idx; 
			int iStart;
			int iEnd; // end of substring index
			if( (idx = str.IndexOf(DOC_CLASS_LIST)) != -1 ) // -1 not found
			{
				if( (iStart = str.IndexOf('=', idx)) != -1 )
				{
					iStart++; // move to next; next of =
					iEnd = str.IndexOfAny( token, iStart );
					_DocumentClassList = str.Substring( iStart, iEnd-iStart );
				}
			}//end of if
		}// end of ExtractDocumentClassList

		/// <summary>
		/// Pass in a line of HTML search page and parse the Repository ID
		/// and predefined repository name.
		/// eg:<OPTION Value=Xs3fhVl-T9zmi3tdrSvKWw==>testdomain1.engmanager3-1.repository
		/// 1) <OPTION Value=Xs3fhVl-T9zmi3tdrSvKWw==>: this is the hash code need to extract.
		/// 2) testdomain1.engmanager3-1.repository: this is the input search name.
		/// 3) Then use equal sign as delima to extract the hash code
		/// </summary>
		/// <param name="strHTML">Input HTML line</param>
		/// <param name="strSearch">Input search String</param>
		public void ExtractRepositoryID(string strHTML, string strSearch)
		{
			int idx; 
			int iStart;
			int iEnd; // end of substring index
			if( (idx = strHTML.IndexOf(strSearch)) != -1 ) // -1 not found
			{
				if( (iEnd = strHTML.LastIndexOf('=',idx)) != -1 ) // reset start from first poisition
				{
					iEnd -= 2; // move just before "==>" ie: point to w
					iStart = strHTML.LastIndexOf( "=", iEnd ); // point to next last = sign
					iStart++; // pos = just after equal sign. ie:point to X
					iEnd++;   // 
					_RepositoryID = strHTML.Substring( iStart, iEnd-iStart ) + "%3D%3D";
					_RepositoryID = _RepositoryID.Trim('"');
					Debug.WriteLine("\nRepository ID = " + _RepositoryID );					
				}// end of if - inner
			}//end of if - outer
		}//end of ExtractRepositoryID

        /// <summary>
        /// Construct Searching URL with search value.
        /// </summary>
        /// <returns>string search URL</returns>
        public string BuildSearchURL( )
        {
            // Construct the URL <hardcode>... this is very long one.
            string strURL  = _BaseURL      						 // https://zantaz.digitalsafe.net/

                + "/zantaz/DigitalSafe/QueryTypeSelectorServlet?"// /zantaz/DigitalSafe/QueryTypeSelectorServlet?
                + SESSION_KEY + getSessionKey()				     // SessionKey = KTYFrFKieoiZ7FoIMvV8dVFSDH7Hc7xZkwLQrmzF3_M=
                + "&"
                + NEW_SEARCH + getNewSearch()		             // NewSearch = new
                + "&"
                + DOC_CLASS_LIST + getDocumentClassList()        // DocumentClassList= D0000001
                + "&"
                + REPOSITORY_ID + getRepositoryID()		         // repositoryID=vrp7IDlwpHXKij7cpuPbIQ%3D%3D
                + "&"
                + DOC_CLASS_LIST2 + getDocumentClassList()       // documentClassList=D0000001
                + "&"
                + U_CONTENT_SEARCH + getUnContentSearch()        // UnsupportedContentSearch=
                + "&"
                + "D_" + FROM + getFrom()						 // D_FROM=
                + "&"
                + FROM + getFrom()								 // FROM=
                + "&"
                + "D_" + APPARENTLY_TO + getTo()				 // D_APPARENTLY-TO=
                + "&"
                + APPARENTLY_TO	+ getTo()					     // APPARENTLY-TO=
                + "&"
                + "D_" + SUBJECT + getSubject()					 // D_SUBJECT=
                + "&"
                + SUBJECT+ getSubject()							 // SUBJECT=
                + "&"
                + "D_" + BODY + getBody()						 // D_%3DBODY%3D=
                + "&"
                + BODY + getBody()								 // %3DBODY%3D=
                + "&"
                + "D_" + ATTACH_NAME+ getAttachName()			 // D_%3DATTACHMENTNAME%3D=
                + "&"
                + ATTACH_NAME+ getAttachName()					 // %3DATTACHMENTNAME%3D=
                + "&"
                + "D_" + MESSAGE_ID + getMessageID()             // D_MESSAGE-ID=
                + "&"
                + MESSAGE_ID + getMessageID()                    // MESSAGE-ID=
                + "&"
                + "D_" + X_ZANTAZ + getXZantaz()                 // D_X-ZANTAZ=
                + "&"
                + X_ZANTAZ + getXZantaz()                        // X-ZANTAZ=
                + "&"
                + AMO + getStartMonth()                          // amo=
                + "&"
                + ADY + getStartDay()                            // ady=
                + "&"
                + AYR + getStartYear()                           // ayr=
                + "&"
                + AHOUR + getStartHour()                         // ahour=
                + "&"
                + AMIN + getStartMinute()                        // amin=
                + "&"
                + BMO + getEndMonth()                            // bmo=
                + "&"
                + BDY + getEndDay()                              // bdy=
                + "&"
                + BYR + getEndYear()                             // byr=
                + "&"
                + BHOUR + getEndHour()                           // bhour=
                + "&"
                + BMIN + getEndMinute()                          // bmin=
                + "&"
                + TZ + getTimeZone()                             // tz (Time Zone)
                ; // end of strURL

            return( strURL );
        }// end of BuildSearchURL

		/// <summary>
		/// Construct the URL that request the search result after the "query page"
		/// </summary>
		/// <returns>request search result URL</returns>
		public string BuildResultPageURL()
		{
            // Construct the URL <hardcode>... this is very long one.
            string strURL  = _BaseURL      				// https://zantaz.digitalsafe.net/
                + "/zantaz/DigitalSafe/SearchServlet?"	// /zantaz/DigitalSafe/SearchServlet?
                + SESSION_KEY + getSessionKey();		// SessionKey = KTYFrFKieoiZ7FoIMvV8dVFSDH7Hc7xZkwLQrmzF3_M=

			return(strURL);
		}// end of BuildResultPageURL

		/// <summary>
		/// Construct the message display URL
		/// 1) Pass in the found HTML line
		/// 2) Filter out the URL parameters
		/// 3) Build the url
		/// </summary>
		/// <param name="str"> HTML string that need to extract info</param>
		/// <returns>file URL string</returns>
		public string BuildMsgDisplayURL( string str )
		{
			string subLine = "";
			int idx = str.IndexOf("showPreviewPage(");
			int iStart = str.IndexOf('(', idx) + 1; // move to next char == '
			int iEnd = str.IndexOf( ')', iStart );
			subLine = str.Substring(iStart, iEnd-iStart);

			string [] splitStr = subLine.Split( new Char [] {','} );

			// Fill url object
			_Navigating = splitStr[1].Trim('\'');
			_QueryStatus = splitStr[2].Trim('\'');
			_Domain = splitStr[3].Trim('\'');
			_SmgID = splitStr[4].Trim('\'');
			_MsgPath = splitStr[5].Trim('\'');
			_Offset = splitStr[6].Trim('\'');
			_ThumbIndex = splitStr[7].Trim('\'');
			_ShowDetails = splitStr[8].Trim('\'');
			_UnparseDate = splitStr[9].Trim('\'');
			_IsUnidxDoc = splitStr[10].Trim('\''); 

			if( _IsUnidxDoc == "true" )
			{
				_FormatDate = splitStr[11].Trim('\'');
				_UnIdxSubject = splitStr[12].Trim('\'');
				_SelectMsgs = splitStr[13].Trim('\'');
				_All = splitStr[14].Trim('\''); 
			}
			else
			{
				_SelectMsgs = splitStr[11].Trim('\'');
				_All = splitStr[12].Trim('\''); 
			}//end of else - end of fill url object data

			// construct the url
			string strURL = _BaseURL						// https://zantaz.digitalsafe.net
				+ "/zantaz/DigitalSafe/MsgDisplayServlet?"	// /zantaz/DigitalSafe/MsgDisplayServlet?
				+ SESSION_KEY + _SessionKey
				+ "&"
				+ NAVIGATING + _Navigating
				+ "&"
				+ QUERY_STATUS + _QueryStatus
				+ "&"
				+ DOMAIN + _Domain
				+ "&"
				+ SMG_ID + _SmgID
				+ "&"
				+ MSG_PATH + _MsgPath
				+ "&"
				+ OFFSET + _Offset
				+ "&"
				+ THUMB_INDEX + _ThumbIndex
				+ "&"
				+ SHOW_DETAILS + _ShowDetails
				+ "&"
				+ UNPARSE_DATE + _UnparseDate
				+ "&"
				+ IS_UNIDX_DOC + _IsUnidxDoc;

			if( _IsUnidxDoc == "true" )
			{
				strURL += "&" 
					+ FORMAT_DATE + _FormatDate
					+ "&"
					+ UNIDX_SUBJECT + _UnIdxSubject;
			}//end of if 

			strURL += "&"
				+ SEL_MSGS + _SelectMsgs
				+ ALL + _All;

			return( strURL );
		}//end of BuildMsgDisplayURL

		/// <summary>
		/// Construct the search query url
		/// Format: https://zantaz.digitalsafe.net/zantaz/DigitalSafe/SimpleSearchFormServlet?DocumentClassList=D0000001&SessionKey=KTYFrFKieoiZ7FoIMvV8dVFSDH7Hc7xZkwLQrmzF3_M=
		/// </summary>
		/// <returns></returns>
		public string BuildSearchFormURL()
		{
			// /DigitalSafe/SimpleSearchFormServlet
			string strURL  = _BaseURL							// https://zantaz.digitalsafe.net
				+ "/zantaz/DigitalSafe/SimpleSearchFormServlet?"// /zantaz/DigitalSafe/SimpleSearchFormServlet?
				+ URLDataObj.DOC_CLASS_LIST						// DocumentClassList=
				+ _DocumentClassList							// D0000001
				+ "&"											// &
				+ URLDataObj.SESSION_KEY						// SessionKey=
				+ _SessionKey;									// KTYFrFKieoiZ7FoIMvV8dVFSDH7Hc7xZkwLQrmzF3_M=

			return(strURL);
		}// end of BuildSearchFormURL

		/// <summary>
		/// eg) https://zantaz.digitalsafe.net/zantaz/DigitalSafe/Logout.html?SessionKey=ExUv8XGXeBW3P5FD42Qic1tNMp91ysVC_p-EgLlcaEU=
		/// </summary>
		/// <returns></returns>
		public string BuildLogoffURL()
		{
			string strURL = _BaseURL					// https://zantaz.digitalsafe.net
				+ "/zantaz/DigitalSafe/Logout.html?"	// /zantaz/DigitalSafe/Logout.html?
				+ URLDataObj.SESSION_KEY				// SessionKey=
				+ _SessionKey;							// ExUv8XGXeBW3P5FD42Qic1tNMp91ysVC_p-EgLlcaEU=
			return( strURL );
		}
	}
}
