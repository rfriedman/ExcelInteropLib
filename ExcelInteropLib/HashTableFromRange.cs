using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class HashTableFromRange: ExcelSheetWorker
    {



        public Hashtable table
        {
            get { return m_htRange; }
        }
        public HashTableFromRange(string fName, string range, int ofset)
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

            CreateTable();

            m_oWB.Close(true, misValue, misValue);
            m_oXL.Quit();

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

        ~HashTableFromRange()
        {
            releaseObject(m_oXL);
            releaseObject(m_oWB);
            releaseObject(m_oSheet);
            releaseObject(m_oRng);

            System.Diagnostics.Trace.WriteLine("HashTableFromRange's destructor is called.");
        }
    }
}
