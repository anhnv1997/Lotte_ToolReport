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
    public partial class frmMail : Form
    {
        private string ID;
        public frmMail(string _ID)
        {
            InitializeComponent();
            this.ID = _ID;
        }

        private void frmMail_Load(object sender, EventArgs e)
        {
            if (this.ID != "")
            {
                txtMailAddr.Text = StaticPool.receiveMails[Convert.ToInt32(this.ID)-1];
                txtID.Text = this.ID;
            }
            else
            {
                if (StaticPool.receiveMails == null)
                {
                    StaticPool.receiveMails = new System.Collections.Specialized.StringCollection();
                    txtID.Text = "1";
                }
                else
                {
                    txtID.Text = StaticPool.receiveMails.Count + 1 + "";
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (this.ID != "")
            {
                StaticPool.receiveMails.RemoveAt(Convert.ToInt32(this.ID)-1);
            }
            StaticPool.receiveMails.Add(txtMailAddr.Text);            
            Properties.Settings.Default.ListMailReceiveReport=StaticPool.receiveMails;
            Properties.Settings.Default.Save();
            this.DialogResult = DialogResult.OK;
        }
    }
}
