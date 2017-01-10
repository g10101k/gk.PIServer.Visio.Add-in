
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
using System.Threading.Tasks;

namespace InduSoft.Visio.Addin
{
    public delegate void ExampleCallback(string txt);

    public partial class ThisAddIn
    {
        private rootRibbon ribbon;
        private log log = new log("ribbon");
        iWorker w;

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
            ribbon.btnWorkClick += ribbon_btnWorkClick;
            ribbon.TimeChange += ribbon_editTimeChanged;
            ribbon.PeriodInSecondsChange += ribbon_PeriodInSecondsChange;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }

        private void ribbon_PeriodInSecondsChange()
        {
            w.Period = ribbon.PeriodInSeconds;
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

        private void ribbon_editTimeChanged()
        {
            w.time = ribbon.TimeText;
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
        public string time = "*";
        public int Period = 15;
        private ExampleCallback callback;
        private Microsoft.Office.Interop.Visio.Page page;
        private log log;
        PISDK.PISDK sdk = new PISDK.PISDK();

        public iWorker(Microsoft.Office.Interop.Visio.Page _page) //Конструктор получает имя функции и номер до кторого ведется счет
        {

            thread = new Thread(new ThreadStart(this.func));
            page = _page;

            //подключенчение к источникам данных:

        }



        public void func()//Функция потока, передаем параметр
        {
            log = new log("iWorker");
            log.Show();
            //свойство "тег": Microsoft.Office.Interop.Visio.Cell cc = vSh.Cells["Prop.Row_1014"]; 

            while (1 == 1)
            {
                if (page != null)
                {
                    foreach (Microsoft.Office.Interop.Visio.Shape vSh in page.Shapes)
                    {
                        //ищем шейпы со значениями и проверяем на группировку
                        CheckGroupShapes(vSh);
                    }
                }
                WaitNSeconds(Period);
            }
        }
        private void WaitNSeconds(int segundos)
        {
            if (segundos < 1) return;
            DateTime _desired = DateTime.Now.AddSeconds(segundos);
            while (DateTime.Now < _desired)
            {
                System.Windows.Forms.Application.DoEvents();
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
                    Microsoft.Office.Interop.Visio.Cell format = vSh.Cells["Prop.Row_14"];
                    Microsoft.Office.Interop.Visio.Cell path = vSh.Cells["Prop.Row_1014"];
                    string[] tmp = path.Formula.Replace("\"","").Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    log.WriteDebug(tmp[0]);

                    PIValue value = GetRTDBData(tmp[0], tmp[1], tmp[2], DateTime.Now);
                    {
                        if (value != null)
                            vSh.Text = value.Value.ToString(format.Formula.Replace("\"", ""));
                    }                    
                }
            }
        }

        public PIValue GetRTDBData(String typeRTDB, String serverName, String tagName, DateTime tM)
        {
            PIValue val = new PIValue();
            switch (typeRTDB)
            {
                #region PI System
                case "PI.":
                    Server piSever = null;
                    try { piSever = sdk.Servers[serverName]; } catch { }
                    if (piSever != null)
                    {
                        if (!piSever.Connected)
                        {
                            piSever.Open();                            
                        }
                        if (piSever.Connected)
                            val = piSever.PIPoints[tagName].Data.ArcValue(time, RetrievalTypeConstants.rtAuto);
                    }
                    break;
                case "AF.":

                    break;
                #endregion
                #region Historian
                case "?Historian":

                    break;
                #endregion
                #region TSDB
                case "?TSDB":

                    break;
                    #endregion
            }
            return val;
        }
    }
}
