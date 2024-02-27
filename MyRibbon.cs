using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TestRibbonWorkbookTracking
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        class WorkbookInfo
        {
            public string Note;
            public bool Flag;
        }

        Application _app;
        IRibbonUI _ribbon;
        bool _shuttingDown;

        Dictionary<Workbook, WorkbookInfo> _workbookInfo = new Dictionary<Workbook, WorkbookInfo>();

        public override string GetCustomUI(string RibbonID)
        {
            return RibbonResources.Ribbon;
        }

        public override object LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = (Application)Application;
        }

        public override void OnBeginShutdown(ref Array custom)
        {
            _shuttingDown = true;
            // Excel seems to get unhappy with some events during shutdown
            _app.WorkbookActivate -= Application_WorkbookActivate;
            _app.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            _app.WorkbookActivate += Application_WorkbookActivate;
            _app.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }

        void Application_WorkbookActivate(Workbook Wb)
        {
            if (!_workbookInfo.ContainsKey(Wb))
                _workbookInfo.Add(Wb, new WorkbookInfo { Note = "<New>", Flag = false });

            _ribbon.Invalidate(); // Ensures the callbacks are called after activate
            Debug.Print("WorkbookActivate: " + Wb.Name);
        }

        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            // NOTE: This is not the best way to track workbook closing, since user can cancel the close.
            //       Tracking all open workbooks in Excel is hard - see e.g. https://stackoverflow.com/questions/58643395/how-to-detect-when-a-workbook-is-closing
            _workbookInfo.Remove(Wb);
        }

        WorkbookInfo GetActiveInfo()
        {
            var wb = _app.ActiveWorkbook;
            if (wb == null)
                return null;
            return _workbookInfo[wb];
        }

        public string GetWorkbookNote(IRibbonControl control)
        {
            var info = GetActiveInfo();
            if (info == null)
                return "<No Wb>";
            return info.Note;
        }

        public void SetWorkbookNote(IRibbonControl control, string text)
        {
            var info = GetActiveInfo();
            if (info == null) 
                return;
            info.Note = text;
        }

        public bool GetWorkbookFlag(IRibbonControl control)
        {
            var info = GetActiveInfo();
            if (info == null) 
                return false;
            return info.Flag;
        }

        public void SetWorkbookFlag(IRibbonControl control, bool value)
        {
            var info = GetActiveInfo();
            if (info == null) 
                return;
            info.Flag = value;
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Hello!");
        }
    }
}
