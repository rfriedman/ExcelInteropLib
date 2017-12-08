using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInteropLib;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;


namespace ExcelLib
{
    public class ExcelSheetWorker : ExcelInteropLib.IExcelSheetWorker
    {
        protected Excel.Application m_oXL;
        protected Excel._Workbook m_oWB;
        protected Excel._Worksheet m_oSheet;
        protected Excel.Range m_oRng;

        protected Hashtable m_htRange;
        protected String[] m_strRange;
        protected int m_iOffset;
        protected object misValue;
        protected Excel.Range m_formatRange;


        public ExcelSheetWorker(string fName, string range, int ofset)
        {
            misValue = System.Reflection.Missing.Value;
            //Start Excel and get Application object.
            m_htRange = new Hashtable();
            m_oXL = new Excel.Application();

            m_oXL.Visible = true;

            //Get a new workbook.
            m_oWB = m_oXL.Workbooks.Open(fName);

            m_oSheet = (Excel.Worksheet)m_oWB.Worksheets.get_Item(1);

            m_strRange = range.Split(',');
            m_iOffset = ofset;
        }

        public void CreateTable()
        {
            string str;
            int rCnt, rw;

            m_oRng = m_oSheet.get_Range(m_strRange[0], m_strRange[1]);
            rw = m_oRng.Rows.Count;

            for (rCnt = 1; rCnt < rw; rCnt++)
            {
                str = (string)(m_oRng.Cells[rCnt, 1] as Excel.Range).Value2;
                if (m_htRange.ContainsKey(str) == false && str.Length > 0)
                {
                    m_htRange.Add(
                        (string)(m_oRng.Cells[rCnt, 1] as Excel.Range).Value2,
                        (double)(m_oRng.Cells[rCnt, m_iOffset] as Excel.Range).Value2);
                }
            }
        }

        public void UpdateRange(Hashtable source)
        {
            string str;
            int rCnt, rw;
            double cnt;
            m_oRng = m_oSheet.get_Range(m_strRange[0], m_strRange[1]);

            rw = m_oRng.Rows.Count;
            str = (string)(m_oRng.Cells[10, 1] as Excel.Range).Value2;
            cnt = (m_oRng.Cells[10, 3] as Excel.Range).Value2;
            for (rCnt = 1; rCnt < rw; rCnt++)
            {
                str = (string)(m_oRng.Cells[rCnt, 1] as Excel.Range).Value2;
                cnt = (double)(m_oRng.Cells[rCnt, m_iOffset] as Excel.Range).Value2;
                if (source.ContainsKey(str) == true && cnt != Double.Parse(source[str].ToString()))
                {
                    m_formatRange = (m_oRng.Cells[rCnt, m_iOffset] as Excel.Range);
                    m_formatRange.Interior.Color = System.Drawing.
                        ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    m_formatRange.Value2 = Double.Parse(source[str].ToString());

                }
            }

            m_oWB.Close(true, misValue, misValue);
            m_oXL.Quit();
        }
         
        void IExcelSheetWorker.ReleaseObject(object obj)
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
