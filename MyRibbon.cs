using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TestRibbonWorkbookTracking
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        IRibbonUI _ribbon;
        Workbook _activeWorkbook;

        public override string GetCustomUI(string RibbonID)
        {
            return RibbonResources.Ribbon;
        }

        public override object LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            var app = (Application)ExcelDnaUtil.Application;
            _activeWorkbook = app.ActiveWorkbook;
            // Track workbook changes with inline lambda event handlers
            app.WorkbookActivate += (wb) =>
            {
                _activeWorkbook = wb;
                _ribbon.Invalidate();
            };
            app.SheetChange += (sh, target) =>
            {
                _ribbon.InvalidateControl("label2");
            };
        }

        public string GetActiveWorkbookLabel(IRibbonControl control)
        {
            return _activeWorkbook?.Name ?? "<No workbook>";
        }

        public string GetA1Label(IRibbonControl control)
        {
            return _activeWorkbook?.ActiveSheet?.Range["A1"].Value2?.ToString() ?? "<No A1>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
    }
}
