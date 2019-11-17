using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace RibbonWorkaround
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RibbonWorkaround.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.ribbon = this.ribbon; // ansteg note - we added this line so that we can get access to the IRibbonUI object from outside this class.
        }

        public string GetWorkbookName(Office.IRibbonControl control)
        {
            // This callback is hooked up to the "Workbook name" textbox on the custom ribbon.
            // Excel will take care of firing it at the right time, and you can use the "control.Context"
            // object to get all the necessary information about the current workbook.
            Excel.Window window = control.Context;
            if (window != null)
            {
                Excel.Worksheet sheet = window.ActiveSheet;
                Excel.Workbook book = sheet.Parent;
                return book.Name;
            } else
            {
                // When the application is booted for the first time this callback will be invoked, but control.Context will be null.
                // In this case, make the control's value blank (it shouldn't matter, because the user can't see the home tab at this point).
                return "";
            }
            
        }

        public bool GetIsConfidential(Office.IRibbonControl control)
        {
            Excel.Window window = control.Context;
            if (window != null)
            {
                Excel.Worksheet sheet = window.ActiveSheet;
                String strVal = (sheet.Cells[1, 1].Value2 ?? "").ToString();
                return strVal.Equals("1");
            }
            else
            {
                // When the application is booted for the first time this callback will be invoked, but control.Context will be null.
                // In this case, make the control's value blank (it shouldn't matter, because the user can't see the home tab at this point).
                return false;
            }
            
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
