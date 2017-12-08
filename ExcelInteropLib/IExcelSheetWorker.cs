using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelInteropLib;

namespace ExcelInteropLib
{
 public interface IExcelSheetWorker
    {
        void UpdateRange(Hashtable source);
        void CreateTable();
        void ReleaseObject(object obj);

    }
}
