using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DBdddhaha
{
    public partial class frmVatTu : Form
    {
        DataAccess.MSACCESS.common cn = new DataAccess.MSACCESS.common();
        public frmVatTu()
        {
            InitializeComponent();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void LoadData()
        {
            DataTable dt = cn.getDataTable("select * from VAT_TU");
            
            dataGridView1.DataSource = dt;
            bindingSource1.DataSource = dt;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource != null)
            {
                if (MessageBox.Show("Có chắn chắn muốn cập nhật?","Hỏi ý kiến",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    DataTable dt = (DataTable)(dataGridView1.DataSource);
                    cn.UpdateDataset(dt, "select * from VAT_TU");
                    MessageBox.Show("Cập nhật thành công","Thông báo");
                    LoadData();
                }
            }
        }

        private void frmVatTu_Load(object sender, EventArgs e)
        {

        }
    }
}
