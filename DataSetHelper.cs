using System;
using System.Data;

namespace QATool
{
	/// <summary>
	/// Summary description for DataSetHelper.
	/// </summary>
	public class DataSetHelper
	{
        public DataSet ds;
		public DataSetHelper( ref DataSet dataSet )
		{
            ds = dataSet;
		}
        public DataSetHelper()
        {
            ds = null;
        }//end of constructor

        /// <summary>
        /// Private method that don't suppose to expose to other
        /// </summary>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        private bool ColumnEqual(object A, object B)
        {

            // Compares two values to see if they are equal. Also compares DBNULL.Value.
            // Note: If your DataTable contains object fields, then you must extend this
            // function to handle them in a meaningful way if you intend to group on them.

            if( A == DBNull.Value && B == DBNull.Value ) //  both are DBNull.Value
                return( true );
            if( A == DBNull.Value || B == DBNull.Value ) //  only one is DBNull.Value
                return( false );

            return( A.Equals(B) );  // value type standard comparison
        }//end of ColumnEqual

        /// <summary>
        /// Return a data table that store the distinct result
        /// </summary>
        /// <param name="TableName">New Table name that store the distinct result</param>
        /// <param name="SourceTable">Source Table</param>
        /// <param name="FieldName">Column header if siyrce table</param>
        /// <returns>Data Table that store the distinct data</returns>
        public DataTable SelectDistinct(string TableName, DataTable SourceTable, string FieldName)
        {
            DataTable dt = new DataTable(TableName);
            dt.Columns.Add(FieldName, SourceTable.Columns[FieldName].DataType);

            object LastValue = null;
            foreach( DataRow dr in SourceTable.Select("", FieldName) )
            {
                if(  LastValue == null || !(ColumnEqual(LastValue, dr[FieldName])) )
                {
                    LastValue = dr[FieldName];
                    dt.Rows.Add(new object[]{LastValue});
                }
            }//end of foreach

            if( ds != null )
                ds.Tables.Add(dt);

            return( dt );
        }//end of SelectDistinct
	}//end of class - DataSetHelper
}//end of namespace - QATool
