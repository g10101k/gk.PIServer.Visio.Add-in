using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;


namespace InduSoft.Visio.Addin
{
    public partial class log : Form
    {
        private System.Windows.Forms.TextBox box;
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ContextMenuStrip contextMenu;
        private System.Windows.Forms.ToolStripMenuItem saveToFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearToolStripMenuItem;

        public log()
        {
            this.components = new System.ComponentModel.Container();
            this.box = new System.Windows.Forms.TextBox();
            this.contextMenu = new System.Windows.Forms.ContextMenuStrip();
            this.saveToFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();

            this.SuspendLayout();

            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.logFormClosing);
            this.Controls.Add(this.box);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
            this.contextMenu.SuspendLayout();
            this.contextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { this.saveToFileToolStripMenuItem, this.clearToolStripMenuItem });
            this.contextMenu.Name = "contMenu";
            this.contextMenu.Size = new System.Drawing.Size(142, 142);

            this.box.BackColor = System.Drawing.Color.Black;
            this.box.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.box.Dock = System.Windows.Forms.DockStyle.Fill;
            this.box.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.box.ForeColor = System.Drawing.Color.White;
            this.box.Multiline = true;
            this.box.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.box.ReadOnly = true;
            this.box.ContextMenuStrip = contextMenu;

            this.saveToFileToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.saveToFileToolStripMenuItem.Text = "&Save To File..";
            this.saveToFileToolStripMenuItem.Click += new System.EventHandler(this.saveToFileToolStripMenuItem_Click);

            this.clearToolStripMenuItem.Size = new System.Drawing.Size(141, 22);
            this.clearToolStripMenuItem.Text = "&Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.clearToolStripMenuItem_Click);
            this.contextMenu.ResumeLayout(false);
            this.ClientSize = new System.Drawing.Size(676, 342);
            this.ShowInTaskbar = false;
            this.PerformLayout();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public void WriteInfo(object msg)
        {
            box.AppendText(string.Format("{0};  [INFO ];    {1}\r\n", DateTime.Now.ToString("o"), msg));
        }
        private void logFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
        }
        public void WriteDebug(object msg)
        {
            box.AppendText(string.Format("{0};  [DEBUG];    {1}\r\n", DateTime.Now.ToString("o"), msg));
        }

        
        public void WriteError(Exception exception, object msg)
        {
            string exMessage = "", exData = "", exStackTrace = "";
            if (exception != null)
            {
                exMessage = exception.Message;
                exData = exception.Data.ToString();
                exStackTrace = exception.StackTrace;
            }

            box.AppendText(string.Format("{0};  [ERROR];    {1}\r\n\t{2}\r\n\t{3}\r\n\t{4}\r\n", DateTime.Now.ToString("o"), msg, exMessage, exData, exStackTrace));
        }
        private void saveToFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
            saveDialog.DefaultExt = "log";
            saveDialog.FileName = "test.log";
            saveDialog.RestoreDirectory = true;
            saveDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*";
            DialogResult dResult = saveDialog.ShowDialog();

            if (dResult == DialogResult.OK)
            {
                File.WriteAllText(saveDialog.FileName, box.Text);
            }
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            box.Clear();
        }
    }
}