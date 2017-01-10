
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
using System.Threading;

namespace InduSoft.Visio.Addin
{
    public delegate void ExampleCallback(string txt);

    public partial class ThisAddIn
    {
        private rootRibbon ribbon;
        private log log = new log();
        private string str = "";
        delegate void SetTextCallbackFromThread(string text);
        iWorker w;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // TODO: Сделать поток который бы опрашивал все сервера
            log.Show();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // TODO: Убить поток который опрашивает сервера
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new rootRibbon();
            ribbon.btnWorkClick += ribbon_btnWorkClick;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }

        private void ribbon_btnTestClicked()
        {
          //  PISDK.PISDK sdk = new PISDK.PISDK();
           // Server ser = sdk.Servers.DefaultServer;
          //  ser.Open();
          //  log.WriteDebug(ser.PIPoints["sinusoid"].Data.Snapshot.Value);
        }

        private void ribbon_btnWorkClick()
        {
            if (ribbon.btnWorkClicked)
            {
                w = new iWorker(this.Application.ActivePage);
                w.thread.Start();
            }
            else if (w != null)
            {
                if (w.thread.IsAlive)
                   w.thread.Suspend();
            }            
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

    class iWorker
    {
        public Thread thread;
        private ExampleCallback callback;
        private Microsoft.Office.Interop.Visio.Page page;
        

        public iWorker(Microsoft.Office.Interop.Visio.Page _page) //Конструктор получает имя функции и номер до кторого ведется счет
        {
            thread = new Thread(new ThreadStart(this.func));
            page = _page;

            

            //подключенчение к источникам данных:
            #region PI SDK
            PISDK.PISDK sdk = new PISDK.PISDK();
            Server piSever = sdk.Servers.DefaultServer;
            piSever.Open();
            if (piSever.Connected) {  }
            #endregion
        }



        public void func()//Функция потока, передаем параметр
        {
            //свойство "тег": Microsoft.Office.Interop.Visio.Cell cc = vSh.Cells["Prop.Row_1014"]; 
  
            foreach (Microsoft.Office.Interop.Visio.Shape vSh in page.Shapes)
            {
                //ищем  шейпы со значениями и проверяем на группировку
                CheckGroupShapes(vSh);
            }
        }

        public void CheckGroupShapes(Microsoft.Office.Interop.Visio.Shape vSh)
        {
            if (vSh.Shapes.Count >= 1)
            {
                foreach (Microsoft.Office.Interop.Visio.Shape vGSh in vSh.Shapes)
                {
                    CheckGroupShapes(vGSh);
                }
            }
            else
            {
                if (vSh.Name.Contains("ISPValue"))
                {
                    //сюда ссылку на обработчик тега шейпа
                    vSh.Text = "0,00";
                }
             
            }
        }

        public void GetRTDBData(String typeRTDB, String serverName, String tagName, DateTime tM)
        {
            switch (typeRTDB)
            {
                #region PI System
                case "PI":
                           
                    break;
                #endregion
                #region Historian
                case "Historian":
                    break;
                #endregion
                #region TSDB
                case "TSDB":
                    break;
                    #endregion
            }
        }
    }
}
