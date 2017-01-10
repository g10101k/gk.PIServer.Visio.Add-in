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
        public event Action TimeChange;
        public event Action PeriodInSecondsChange;
        public bool btnWorkClicked { get { return btnWork.Checked; } set { btnWork.Checked = value; } }
        /// <summary>
        /// Время запроса данных в формате PI по умолчанию "*"
        /// </summary>
        public string TimeText { get { return editTime.Text; } set { editTime.Text = value; } }
        /// <summary>
        /// Период обновления данных в секундах, не может быть меньше 5, по умолчанию 15
        /// </summary>
        public int PeriodInSeconds { get { return Convert.ToInt32(editPeriodInSeconds.Text); } set { editPeriodInSeconds.Text = value > 5 ? value.ToString():"15"; } }

        private void InduSoft_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (btnWorkClick != null)
                btnWorkClick();
            if (btnWorkClicked)
                btnWork.Image = GlobalResource.iBug;
            else
                btnWork.Image = GlobalResource.iHammer;


        }

        private void editTime_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (TimeChange != null)
                if (true) { // TODO: проверить на правильность времени
                    TimeChange();
                }
        }

        private void editPeriodInSeconds_TextChanged(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int i = Convert.ToInt32(editPeriodInSeconds.Text);
                if (i < 5)
                    throw new Exception();
                if (PeriodInSecondsChange != null)
                    PeriodInSecondsChange();
            }
            catch {
                editPeriodInSeconds.Text = "15";
            }
            
        }
    }
}
