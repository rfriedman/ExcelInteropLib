using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public abstract class ExcelSheetWorker
    {
        protected Excel.Application m_oXL;
        protected Excel._Workbook m_oWB;
        protected Excel._Worksheet m_oSheet;
        protected Excel.Range m_oRng;

        protected Hashtable m_htRange;
        protected String[] m_strRange;
        protected int m_iOffset;
        protected object misValue;
        

        public ExcelSheetWorker() { }

        protected void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Diagnostics.Trace.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
