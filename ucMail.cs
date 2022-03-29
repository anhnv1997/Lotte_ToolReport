using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ToolReport.Forms;

namespace ToolReport
{
    public partial class ucMail : UserControl
    {
        public ucMail()
        {
            InitializeComponent();
        }

        private void tsbADD_Click(object sender, EventArgs e)
        {
            frmMail frm = new frmMail("");
            frm.Text = "Add mail address";
            frm.ShowDialog();
            if (frm.DialogResult == DialogResult.OK)
            {
                LoadDataGridView();
            }
        }

        private void tsbEDIT_Click(object sender, EventArgs e)
        {
            string ID = StaticPool.Id(dgvMail);
            if (ID != "")
            {
                frmMail frm = new frmMail(ID);
                frm.ShowDialog();
                frm.Text = "Edit Mail Address";
                if (frm.DialogResult == DialogResult.OK)
                {
                    LoadDataGridView();
                }
            }
        }

        private void tsbDELETE_Click(object sender, EventArgs e)
        {
            int ID = Convert.ToInt32(StaticPool.Id(dgvMail));
            if (StaticPool.Id(dgvMail) != "")
            {
                if(MessageBox.Show("Do you want delete this record", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    StaticPool.receiveMails.RemoveAt(ID - 1);
                    Properties.Settings.Default.ListMailReceiveReport = StaticPool.receiveMails;
                    Properties.Settings.Default.Save();
                    LoadDataGridView();
                }
            }
            else
            {
                MessageBox.Show("Please select a record");
            }
        }

        public void LoadDataGridView()
        {
            dgvMail.Rows.Clear();
            foreach(string str in StaticPool.receiveMails)
            {
                dgvMail.Rows.Add(dgvMail.Rows.Count + 1, str);
            }
        }
    }
}
