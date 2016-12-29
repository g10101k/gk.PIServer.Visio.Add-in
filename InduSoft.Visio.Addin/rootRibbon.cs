using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace InduSoft.Visio.Addin
{
    public partial class rootRibbon
    {
        public event Action btnTestClicked;
        public event Action button1ClickEd;
        private void InduSoft_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTest_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnTestClicked != null)
                btnTestClicked();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (button1ClickEd != null)
                button1ClickEd();
        }
    }
}
