using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GPACalc
{
    public partial class ChooseExcelColumns : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        // the datatable that should be choosed
        DataTable dtexcel;

        // the datatable that should be out
        DataTable dtout;

        // choosing index
        int index = 0;

        public ChooseExcelColumns()
        {
            InitializeComponent();
        }

        public void SetData(ref DataTable dt,ref DataTable outdt)
        {
            dtexcel = dt;
            dtout = outdt;
        }

        private void simpleButton_Click(object sender, EventArgs e)
        {
            switch (index)
            {
                case 1:foreach (DataRow dr in dtexcel.Rows)
                {
                    DataRow drn = dtout.NewRow();
                    drn[0] = dr[gridView.FocusedColumn.AbsoluteIndex];
                    dtout.Rows.Add(drn);
                }
                index++;
                labelControl.Text="Student Name";
                break;

                case 2:for (int i=0;i<dtexcel.Rows.Count;i++)
                {
                    dtout.Rows[i][1] = dtexcel.Rows[i][gridView.FocusedColumn.AbsoluteIndex];
                }
                index++;
                labelControl.Text = "Course Name";
                break;

                case 3: for (int i = 0; i < dtexcel.Rows.Count; i++)
                {
                    dtout.Rows[i][2] = dtexcel.Rows[i][gridView.FocusedColumn.AbsoluteIndex];
                }
                index++;
                labelControl.Text = "Credits";
                break;

                case 4: for (int i = 0; i < dtexcel.Rows.Count; i++)
                {
                    dtout.Rows[i][3] = dtexcel.Rows[i][gridView.FocusedColumn.AbsoluteIndex];
                }
                index++;
                labelControl.Text = "Score";
                simpleButton.Text = "Finish";
                break;

                case 5: for (int i = 0; i < dtexcel.Rows.Count; i++)
                {
                    dtout.Rows[i][4] = dtexcel.Rows[i][gridView.FocusedColumn.AbsoluteIndex];
                }
                this.Close();
                break;
            }
        }

        private void ChooseExcelColumns_Load(object sender, EventArgs e)
        {
            gridControl.DataSource = dtexcel;
            gridView.PopulateColumns();
            gridControl.RefreshDataSource();

            labelControl.Text = "Student ID";
            index = 1;
        }
    }
}
