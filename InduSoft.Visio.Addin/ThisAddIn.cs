﻿
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
            ribbon.button1ClickEd += button1_Click;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }

        private void ribbon_btnTestClicked()
        {
            iWorker w = new iWorker(new ExampleCallback(ResultCallback));
            Thread t = new Thread(new ThreadStart(w.func));
            t.Start();
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
        Thread thread;
        private ExampleCallback callback;

        public iWorker(ExampleCallback _callback) //Конструктор получает имя функции и номер до кторого ведется счет
        {
            callback = _callback;
            //thread = new Thread(this.func);
            //thread.Name = name;
            //thread.Start(10);//передача параметра в поток
        }

        public void func()//Функция потока, передаем параметр
        {
            for (int i = 0; i < (int)100; i++)
            {
                callback(i.ToString());
                Thread.Sleep(1000 * new Random().Next(5));
            }
        }
    }
}
