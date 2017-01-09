
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
        private Dictionary<string, Data> vals = new Dictionary<string, Data>();


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
            ribbon.btnTestClicked += ribbon_btnTestClicked;
            ribbon.btnFindISPValueClicked += ribbon_btnFindISPValueClicked;
            ribbon.button1ClickEd += button1_Click;
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

        private void button1_Click()
        {
            log.WriteDebug(str);
        }
        /// <summary>
        /// Через эту функцию поток вернет свой текст
        /// </summary>
        /// <param name="txt">Возвращаемый текст</param>
        public void ResultCallback(string txt)
        {
            str = txt;
            Microsoft.Office.Interop.Visio.Document vD = this.Application.ActiveDocument;
            Microsoft.Office.Interop.Visio.Page vAP = vD.Application.ActivePage;

            foreach (Microsoft.Office.Interop.Visio.Shape vSh in vAP.Shapes)
            {
                if (vSh.Name.Contains("ISPValue"))
                {
                    vSh.Text = txt;
                    try
                    {
                        Microsoft.Office.Interop.Visio.Cell cc = vSh.Cells["Prop.Row_1014"]; //
                        //log.WriteDebug(cc.Formula);
                    }
                    catch(Exception ex)
                    //foreach (Microsoft.Office.Interop.Visio.Cell c in vSh.Cells[1])
                    {
                        //log.WriteError(ex, null);
                    }

                }
            }
        }
      
        public void CheckGroupShapes(Microsoft.Office.Interop.Visio.Shape vSh)
        {
            foreach (Microsoft.Office.Interop.Visio.Shape vGSh in vSh.Shapes)
                {
                    if (vGSh.Name.Contains("ISPValue"))
                    {
                        //сюда ссылку на обработчик тега шейпа
                        vGSh.Text = "0,00";
                    }
                    else if (vGSh.Shapes.Count >= 1) CheckGroupShapes(vGSh);
                }
        }
        private void ribbon_btnFindISPValueClicked()
        {
            Microsoft.Office.Interop.Visio.Document vD = this.Application.ActiveDocument;
            Microsoft.Office.Interop.Visio.Page vAP = vD.Application.ActivePage;
            
            foreach (Microsoft.Office.Interop.Visio.Shape vSh in vAP.Shapes)
            {
                //ищем  шейпы со значениями и проверяем на группировку
                if (vSh.Shapes.Count >= 1)
                {
                    CheckGroupShapes(vSh);
                }
                //если группировки нет 
                else if (vSh.Name.Contains("ISPValue"))
                {
                    // сюда ссылку на обработчик тега шейпа
                    vSh.Text = "0,00";
                    
                }
            }
        }

        private void ribbon_btnWorkClick()
        {
            if (ribbon.btnWorkClicked)
            {
                w = new iWorker(new ExampleCallback(ResultCallback), ref vals);
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
        private Dictionary<string, Data> vals;

        public iWorker(ExampleCallback _callback, ref Dictionary<string, Data> _vals) //Конструктор получает имя функции и номер до кторого ведется счет
        {
            callback = _callback;
            thread = new Thread(new ThreadStart(this.func));
            vals = _vals;
        }

        public void func()//Функция потока, передаем параметр
        {
            if (vals.Count == 0)
            {
                vals.Add("TEST", new Data("path", "", DateTime.Now));
            }
            
            for (int i = 0; i < (int)100; i++)
            {
                callback(i.ToString());
                Thread.Sleep(1000 * new Random().Next(5));
            }
        }
    }

    class Data {
        public string Path { get; set; }
        public string Value { get; set; }
        public DateTime Date { get; set; }

        public Data(string _path, string _value, DateTime _date)
        {
            Path = _path;
            Value = _value;
            Date = _date;
        }
    }
}
