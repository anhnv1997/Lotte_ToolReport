using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolReport.Forms
{
    public partial class frmSetting : Form
    {
        public frmSetting()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.TimeSendReport = dtpTimeSendReport.Value.ToString("yyyy/MM/dd HH:mm:ss");
            Properties.Settings.Default.Save();
            this.DialogResult = DialogResult.OK;
        }

        private void frmSetting_Load(object sender, EventArgs e)
        {
            dtpTimeSendReport.Value = DateTime.Parse(Properties.Settings.Default.TimeSendReport);
            if(StaticPool.receiveMails!=null)
                ucMail1.LoadDataGridView();
        }
    }
}
