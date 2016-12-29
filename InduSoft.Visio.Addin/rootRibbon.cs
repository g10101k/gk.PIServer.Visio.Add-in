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
        private void InduSoft_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTest_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnTestClicked != null)
                btnTestClicked();
        }
    }
}
