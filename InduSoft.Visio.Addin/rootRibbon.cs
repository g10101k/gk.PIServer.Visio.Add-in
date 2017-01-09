using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;

namespace InduSoft.Visio.Addin
{
    public partial class rootRibbon
    {
        public event Action btnWorkClick;
        public bool btnWorkClicked { get { return btnWork.Checked; } set { btnWork.Checked = value; } }

        private void InduSoft_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnWorkClick != null)
                btnWorkClick();
        }
    }
}
