
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Win32;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using PISDK;

namespace InduSoft.Visio.Addin
{
    public partial class ThisAddIn
    {
        private rootRibbon ribbon;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // TODO: Сделать поток который бы опрашивал все сервера
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // TODO: Убить поток который опрашивает сервера
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new rootRibbon();
            ribbon.btnTestClicked += ribbon_btnTestClicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon })
        }

        private void ribbon_btnTestClicked()
        {
            PISDK.PISDK sdk = new PISDK.PISDK();
            Server ser = sdk.Servers.DefaultServer;
            ser.Open();
            log.WriteDebug(ser.PIPoints["sinusoid"].Data.Snapshot.Value);
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
