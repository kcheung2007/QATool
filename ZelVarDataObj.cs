using System;

namespace QATool
{
	/// <summary>
	/// Summary description for ZelVarDataObj.
	/// </summary>
	public class ZelVarDataObj
	{
        private string _run = "default";
        private string _to  = "change@me.com";
        // add new here

		public ZelVarDataObj()
		{
			// TODO: Add constructor logic here
		}// end of constructor

        public string RUN // display text
        {
            get
            {
                return _run;
            }
            set
            {
                _run = value;
            }
        }//end of property - varRun

        public string TO // display text
        {
            get
            {
                return _to;
            }
            set
            {
                _to = value;
            }
        }
	}//end of class - ZelVarDataObj
}