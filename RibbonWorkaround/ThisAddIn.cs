using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace RibbonWorkaround
{
    public partial class ThisAddIn
    {

        public Office.IRibbonUI ribbon; // ansteg note - this is populated in the Ribbon_Load method of Ribbon1. It gives access to the ribbon from other parts of our application (see Application_SheetChange below).

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookOpen += Application_WorkbookOpen;
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += ThisAddIn_NewWorkbook;
            this.Application.SheetChange += Application_SheetChange;
        }

        private void ThisAddIn_NewWorkbook(Excel.Workbook Wb)
        {
            ribbon.Invalidate();
        }

        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            ribbon.Invalidate();
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            ribbon.Invalidate();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
