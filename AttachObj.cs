using System;
using System.IO;


namespace QATool
{
	/// <summary>
	/// Summary description for AttachObj.
	/// </summary>
	public class AttachObj
	{
		private int m_idxAttach = 0; 
		private int m_numFile = 0;
		private string m_inFolder = "";

		private DirectoryInfo m_di;
		private FileInfo[] m_lstFiles;

		public AttachObj( string inFolder )
		{
			m_inFolder = inFolder;			
			m_di = new DirectoryInfo(inFolder); // attachment folder
			m_lstFiles = m_di.GetFiles();
			m_numFile = m_lstFiles.Length;
		}//end of constructor

		public int idxAttach
		{
			get
			{
				return m_idxAttach;
			}
			set
			{
				m_idxAttach = value;
			}
		}// end of idxAttach

		public int numFile
		{
			get
			{
				return( m_numFile );
			}
		}// end of numFile

		public string attchFullName
		{
			get
			{
				return( m_lstFiles[m_idxAttach].FullName );
			}
		}//end of attchFullFame

        public FileInfo[] getListFiles
        {
            get
            {
                return( m_lstFiles );
            }
        }
	}
}
