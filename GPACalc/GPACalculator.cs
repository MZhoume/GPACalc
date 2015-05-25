using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace GPACalc
{
    public partial class GPACalculator : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        // the max size of byte[] when open or save to file
        const int MAXBYTESIZE = 100000;

        // DataTbale that contains all the data that user imported or entered
        DataTable dtCalculatedData;

        // Datatable that only contains overall gpas
        DataTable dtOverallGPAData;

        // datatable that contains custom data
        DataTable dtCustomData;

        // imported excel data
        DataTable dtExcelData;

        // particular student's data
        DataTable dtParCalculatedData;

        // particular student's name
        string strParStuName;

        // Show if the window is in full screen mode
        bool isfullscreen = false;

        // If there is data in this datatable
        bool IsCalculated = false;

        // dataset that contains all the gpa methods
        DataSet dsGPAMethods;

        // Set the dataset that contains the three custom methods to calculate GPA
        DataSet dsCustomGPAMethods;

        public GPACalculator()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Auto change the size of the other components when the window is in full screen mode
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFullScreenButtonClicked(MouseEventArgs e)
        {
            if (isfullscreen)
            {
                xtraTabControl.Height -= 110;
                gridcCustomData.Height -= 110;
                gridcCalculatedData.Height -= 110;
                chartcCompareGPAs.Height -= 110;
                xtraTabControl.Location = new Point(xtraTabControl.Location.X, xtraTabControl.Location.Y + 110);
            }
            else
            {
                xtraTabControl.Location = new Point(xtraTabControl.Location.X, xtraTabControl.Location.Y - 110);
                xtraTabControl.Height += 110;
                gridcCustomData.Height += 110;
                gridcCalculatedData.Height += 110;
                chartcCompareGPAs.Height += 110;
            }
            isfullscreen = !isfullscreen;
            base.OnFullScreenButtonClicked(e);
        }

        /// <summary>
        /// Auto change the size of the other components when ribbon control colapsing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ribbonControl_MinimizedChanged(object sender, EventArgs e)
        {
            if (ribbonControl.Minimized)
            {
                xtraTabControl.Location = new Point(xtraTabControl.Location.X, xtraTabControl.Location.Y - 95);
                xtraTabControl.Height += 95;
                gridcCustomData.Height += 95;
                gridcCalculatedData.Height += 95;
                chartcCompareGPAs.Height += 95;
            }
            else
            {
                xtraTabControl.Height -= 95;
                gridcCustomData.Height -= 95;
                gridcCalculatedData.Height -= 95;
                chartcCompareGPAs.Height -= 95;
                xtraTabControl.Location = new Point(xtraTabControl.Location.X, xtraTabControl.Location.Y + 95);
            }
        }

        private void barbtnitmNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            barstaitmLeftInfo.Caption = "Creating a new workspace";

            // clear all data that used so that all the workspace is brand new
            dtCalculatedData.Clear();
            dtCustomData.Clear();
            dtOverallGPAData.Clear();
            chartcCompareGPAs.Series.Clear();

            baredtitmMyName.EditValue = null;

            IsCalculated = false;
            barchkitmShowOnlyOverallGPA.Checked = false;

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmCustomData_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // select the custom data page
            xtraTabControl.SelectedTabPageIndex = 1;
            return;
        }

        private void barbtnitmPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (IsCalculated == false)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Please Finish Calculation first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            gridcCalculatedData.ShowRibbonPrintPreview();
            return;
        }

        private void barbtnitmContactClive_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Help.ContactClive();
        }

        private void tiContactClive_ItemClick(object sender, DevExpress.XtraEditors.TileItemEventArgs e)
        {
            Help.ContactClive();
        }

        private void barbtnitmbtmContactClive_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Help.ContactClive();
        }

        private void barbtnitmAbout_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ribbonControl.ShowApplicationButtonContentControl();
        }

        private void barbtnitmSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            saveFileDialog.Filter = "GPA Calculator File|*.gpac|All Files|*.*";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            // export all the data that contained in the custom datatable
            dtCustomData.WriteXml(saveFileDialog.FileName);
        }

        private void barbtnitmOpen_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openFileDialog.Filter = "GPA Calc File|*.GPAC|All Files|*.*";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            // import all the data that contained in the custom datatable
            dtCustomData.ReadXml(openFileDialog.FileName);
        }

        private void barbtnitmSettings_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl.SelectedTabPageIndex = 4;
        }

        private void barchkitmShowOnlyOverallGPA_CheckedChanged(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // for this method is firing after the checked changed
            if (!IsCalculated)
            {
                if (barchkitmShowOnlyOverallGPA.Checked==true)
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show("Please Finish Calculation first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    barchkitmShowOnlyOverallGPA.Checked = false;
                }
                return;
            }
            barstaitmLeftInfo.Caption = "Working";

            if (barchkitmShowOnlyOverallGPA.Checked)
            {
                gridcCalculatedData.DataSource = dtOverallGPAData;
                gridvCalculatedData.PopulateColumns();
                gridcCalculatedData.RefreshDataSource();
            }
            else
            {
                gridcCalculatedData.DataSource = dtCalculatedData;
                gridvCalculatedData.PopulateColumns();
                gridcCalculatedData.RefreshDataSource();
            }

            // select the correspond tabpage
            xtraTabControl.SelectedTabPageIndex = 2;

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmStart_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (IsCalculated)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Please create a new workspace!");
                return;
            }

            if (dtCustomData.Rows.Count == 0)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Please insert data first!");
                return;
            }

            strParStuName = baredtitmMyName.EditValue == null ? null : baredtitmMyName.EditValue.ToString();

            if (strParStuName==null)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Please choose your name in the custom data first!");
                return;
            }

            var qParStuNametmp = from dt in dtCustomData.AsEnumerable()
                                 where dt.Field<string>("Student Name") == strParStuName
                                 select dt;

            if (qParStuNametmp.Count<DataRow>() == 0)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Please choose your name in the custom data first!");
                return;
            }

            barstaitmLeftInfo.Caption = "CalCulating";
            //backgroundWorker.DoWork += new DoWorkEventHandler(DoCalculate);
            //backgroundWorker.RunWorkerAsync();

            int index = 0;
            switch (ribgalbaritmCalcMethod.Gallery.GetCheckedItems()[0].Caption)
            {
                case "百分算法": index = 0; break;
                case "5分算法": index = 1; break;
                case "标准算法": index = 2; break;
                case "北大算法": index = 3; break;
                case "浙大算法": index = 4; break;
                case "上交算法": index = 5; break;
                case "中科大算法": index = 6; break;
                case "自定义1": index = 10; break;
                case "自定义2": index = 11; break;
                case "自定义3": index = 12; break;
            }

            if (index < 10)
            {
                Data.CalculteGPAs(ref dtCustomData, ref dtCalculatedData, ref dtOverallGPAData, dsGPAMethods.Tables[index]);
            }
            else
            {
                Data.CalculteGPAs(ref dtCustomData, ref dtCalculatedData, ref dtOverallGPAData, dsCustomGPAMethods.Tables[index - 10]);
            }

            IsCalculated = true;

            // sort the overall gpa
            dtOverallGPAData.DefaultView.Sort = "Overall GPA";

            // get particular student's name
            var qParStuName = from dt in dtCalculatedData.AsEnumerable()
                              where dt.Field<string>("Student Name") == strParStuName
                              select dt;
            dtParCalculatedData = qParStuName.CopyToDataTable<DataRow>();

            // Display the datatable in gridcontrol
            gridcCalculatedData.DataSource = dtCalculatedData;
            gridvCalculatedData.PopulateColumns();
            gridcCalculatedData.RefreshDataSource();

            DevExpress.XtraEditors.XtraMessageBox.Show("Calculate successful!");

            // display the correspond tabpage
            xtraTabControl.SelectedTabPageIndex = 2;
        }

        /*private void DoCalculate(object obj, DoWorkEventArgs e)
        {
            backgroundWorker.DoWork -= new DoWorkEventHandler(DoCalculate);

            int index = 0;
            switch (ribgalbaritmCalcMethod.Gallery.GetCheckedItems()[0].Caption)
            {
                case "百分算法": index = 0; break;
                case "5分算法": index = 1; break;
                case "标准算法": index = 2; break;
                case "北大算法": index = 3; break;
                case "浙大算法": index = 4; break;
                case "上交算法": index = 5; break;
                case "中科大算法": index = 6; break;
                case "自定义1": index = 10; break;
                case "自定义2": index = 11; break;
                case "自定义3": index = 12; break;
            }

            if (index < 10)
            {
                Data.CalculteGPAs(ref dtCustomData, ref dtCalculatedData, ref dtOverallGPAData, dsGPAMethods.Tables[index]);
            }
            else
            {
                Data.CalculteGPAs(ref dtCustomData, ref dtCalculatedData, ref dtOverallGPAData, dsCustomGPAMethods.Tables[index - 10]);
            }

            IsCalculated = true;

            // sort the overall gpa
            dtOverallGPAData.DefaultView.Sort = "Overall GPA";

            // get particular student's name
            var qParStuName=from dt in dtCalculatedData.AsEnumerable()
                            where dt.Field<string>("Student Name")==strParStuName
                            select dt;
            dtParCalculatedData = qParStuName.CopyToDataTable<DataRow>();

            DevExpress.XtraEditors.XtraMessageBox.Show("Calculate successful!");
        }*/

        private void ribgalbaritmCalcMethod_GalleryItemClick(object sender, DevExpress.XtraBars.Ribbon.GalleryItemClickEventArgs e)
        {
            DevExpress.XtraBars.Ribbon.GalleryItem gi = ribgalbaritmCalcMethod.Gallery.GetCheckedItems()[0];
            switch (gi.Caption)
            {
                case "自定义1":
                    // for the table is actually exists but it doesn't contain any data
                    if (dsCustomGPAMethods.Tables[0].Rows.Count==0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("Please define the method 1 first by using the setting page!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    } break;
                case "自定义2":
                    if (dsCustomGPAMethods.Tables[1].Rows.Count==0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("Please define the method 2 first by using the setting page!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    } break;
                case "自定义3":
                    if (dsCustomGPAMethods.Tables[2].Rows.Count==0)
                    {
                        DevExpress.XtraEditors.XtraMessageBox.Show("Please define the method 3 first by using the setting page!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    } break;
            }
        }

        private void GPACalculator_Load(object sender, EventArgs e)
        {
            barstaitmLeftInfo.Caption = "Initializing";

            // create a new dataset contains the gpa methods
            dsGPAMethods = new DataSet() { DataSetName = "GPAC GPA Methods" };

            // if generated
            if (File.Exists(Application.StartupPath + "\\Methods.xml"))
            {
                Data.GenerateGPAMethodsDataSetColumns(ref dsGPAMethods);
                dsGPAMethods.ReadXml(Application.StartupPath + "\\Methods.xml");
            }
            else
            {
                // create a new one and save it
                Data.GenerateCommonGPAMethods(ref dsGPAMethods);
                dsGPAMethods.WriteXml(Application.StartupPath + "\\Methods.xml");
            }

            // create a new dataset contains custom gpa methods
            dsCustomGPAMethods = new DataSet() { DataSetName = "GPAC Custom GPA Methods" };
            Data.GenerateCustomGPAMethodsDataSetColumns(ref dsCustomGPAMethods);

            // if generated
            if (File.Exists(Application.StartupPath + "\\CustomMethods.xml"))
            {
                dsCustomGPAMethods.ReadXml(Application.StartupPath + "\\CustomMethods.xml");
            }

            // Preset user customed GPA methods to setting gridcontrol
            gridcCustomGPAMethods.DataSource = dsCustomGPAMethods.Tables[0];
            gridvCustomGPAMethods.PopulateColumns();
            gridvCustomGPAMethods.Columns[0].ColumnEdit = repositoryItemTextEditNum;
            gridvCustomGPAMethods.Columns[1].ColumnEdit = repositoryItemTextEditNum;
            gridcCustomGPAMethods.RefreshDataSource();
            // check the custom gpa method radio
            radgrpCustomGPAMethods.SelectedIndex = 0;

            // Preset user custome datatable
            dtCustomData = new DataTable() { TableName = "GPAC Custome Data" };
            Data.GenerateCustomeDataTableColumns(ref dtCustomData);
            gridcCustomData.DataSource = dtCustomData;
            gridvCustomData.PopulateColumns();
            gridvCustomData.Columns[0].ColumnEdit = repositoryItemTextEditNum;
            gridvCustomData.Columns[3].ColumnEdit = repositoryItemTextEditNum;
            gridvCustomData.Columns[4].ColumnEdit = repositoryItemTextEditNum;
            gridcCustomData.RefreshDataSource();

            // Preset calculated datatable
            dtCalculatedData = new DataTable() { TableName = "GPAC Calculated Data" };
            Data.GenerateCalculatedDataTableColumns(ref dtCalculatedData);

            // Preset overall gpas datatable
            dtOverallGPAData = new DataTable() { TableName = "GPAC Overall GPA Data" };
            Data.GenerateCalculatedDataTableColumns(ref dtOverallGPAData);
            dtOverallGPAData.Columns.Remove("Course Name");
            dtOverallGPAData.Columns.Remove("Credits");
            dtOverallGPAData.Columns.Remove("Score");
            dtOverallGPAData.Columns.Remove("GPA");
            dtOverallGPAData.Columns.Add("Overall GPA");
            dtOverallGPAData.Columns[2].DataType = Type.GetType("System.Single");

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void radgrpCustomGPAMethods_SelectedIndexChanged(object sender, EventArgs e)
        {
            barstaitmLeftInfo.Caption = "Changing";

            // change the current gpamethods datatable
            gridcCustomGPAMethods.DataSource = dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex];
            gridvCustomGPAMethods.PopulateColumns();
            gridvCustomGPAMethods.Columns[0].ColumnEdit = repositoryItemTextEditNum;
            gridvCustomGPAMethods.Columns[1].ColumnEdit = repositoryItemTextEditNum;
            gridcCustomGPAMethods.RefreshDataSource();

            // display it in the button
            drpdbtnCustomGPAMethods.Text = radgrpCustomGPAMethods.Properties.Items[radgrpCustomGPAMethods.SelectedIndex].Value.ToString();

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void simbtnSaveCustomGPAMethods_Click(object sender, EventArgs e)
        {
            barstaitmLeftInfo.Caption = "Saving";
            backgroundWorker.DoWork += new DoWorkEventHandler(DoSaveCustomGPAMethods);
            backgroundWorker.RunWorkerAsync();
        }

        private void DoSaveCustomGPAMethods(object obj, DoWorkEventArgs e)
        {
            backgroundWorker.DoWork -= new DoWorkEventHandler(DoSaveCustomGPAMethods);

            // check if the datatable is empty
            if (dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex].Rows.Count == 0)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Current Custom Method is empty!");
                return;
            }

            // auto add score 0 when there isn't any
            if (dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex]
                .Rows[dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex].Rows.Count - 1][0].ToString() != "0")
            {
                DataRow dr = dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex].NewRow();
                dr[0] = dr[1] = 0;
                dsCustomGPAMethods.Tables[radgrpCustomGPAMethods.SelectedIndex].Rows.Add(dr);
            }

            // if already have one then replacee it
            if (File.Exists(Application.StartupPath + "\\CustomMethods.xml"))
            {
                File.Delete(Application.StartupPath + "\\CustomMethods.xml");
            }

            // generate customemethods.xml
            dsCustomGPAMethods.WriteXml(Application.StartupPath + "\\CustomMethods.xml");
            DevExpress.XtraEditors.XtraMessageBox.Show("Successful Saved!", "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void gridvCustomData_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {
            // for the e is actually a datarowview
            DataRow dr = (e.Row as DataRowView).Row;
            if (dr[0] == System.DBNull.Value || dr[3] == System.DBNull.Value || dr[4] == System.DBNull.Value)
            {
                e.ErrorText = "Please insert the 'Student ID' , 'Credit' and the 'Score' cell!";
                e.Valid = false;
            }
            else
            {
                e.Valid = true;
            }
        }

        private void gridvCustomGPAMethods_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {
            // for the e is actually a datarowview
            DataRow dr = (e.Row as DataRowView).Row;
            if (dr[0] == System.DBNull.Value || dr[1] == System.DBNull.Value)
            {
                e.ErrorText = "Please insert the 'Score From' and the 'GPA' cell!";
                e.Valid = false;
            }
            else
            {
                e.Valid = true;
            }
        }

        private void barbtnImportExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            openFileDialog.Filter = "Excel File|*.xls|All Files|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            barstaitmLeftInfo.Caption = "Importing";
            try
            {
                dtExcelData = ExcelTool.ExcelToDatatable(openFileDialog.FileName);
            }
            catch
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Import Error! Please Check The Chosen File.\n", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DevExpress.XtraEditors.XtraMessageBox.Show("Import File Successful! Please Choose the correspond Columns.", "Finish", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // open new form to choose from imported data
            ChooseExcelColumns cecf = new ChooseExcelColumns();
            cecf.SetData(ref dtExcelData, ref dtCustomData);
            cecf.ShowDialog();

            xtraTabControl.SelectedTabPageIndex = 1;
            barstaitmLeftInfo.Caption = "Ready";
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmHelp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl.SelectedTabPageIndex = 5;
        }

        private void barbtnExportExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            saveFileDialog.Filter = "Excel WorkTable File|*.xls|All Files|*.*";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            barstaitmLeftInfo.Caption = "Exporting";

            gridcCalculatedData.ExportToXls(saveFileDialog.FileName);

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmCompGPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            chartcCompareGPAs.Series.Clear();

            DevExpress.XtraCharts.Series series = new DevExpress.XtraCharts.Series("Overall GPA", DevExpress.XtraCharts.ViewType.Bar);

            // set the series
            series.DataSource = dtOverallGPAData;
            series.ArgumentScaleType = DevExpress.XtraCharts.ScaleType.Qualitative;
            series.ArgumentDataMember = "Student Name";
            series.ValueScaleType = DevExpress.XtraCharts.ScaleType.Numerical;
            series.ValueDataMembers.AddRange(new string[] { "Overall GPA" });

            chartcCompareGPAs.Series.Add(series);

            xtraTabControl.SelectedTabPageIndex = 3;
        }

        private void barbtnitmAnalyseGPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            chartcCompareGPAs.Series.Clear();

            DevExpress.XtraCharts.Series series = new DevExpress.XtraCharts.Series("GPA", DevExpress.XtraCharts.ViewType.Bar);

            // set the series
            series.DataSource = dtParCalculatedData;
            series.ArgumentScaleType = DevExpress.XtraCharts.ScaleType.Qualitative;
            series.ArgumentDataMember = "Course Name";
            series.ValueScaleType = DevExpress.XtraCharts.ScaleType.Numerical;
            series.ValueDataMembers.AddRange(new string[] { "GPA" });

            chartcCompareGPAs.Series.Add(series);

            xtraTabControl.SelectedTabPageIndex = 3;
        }

        private void barbtnitmSortGPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            dtParCalculatedData.DefaultView.Sort = "GPA";
            barbtnitmAnalyseGPA_ItemClick(sender, e);
        }

        private void barbtnitmsubECompGPAs_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            barbtnitmCompGPA_ItemClick(sender, e);

            saveFileDialog.Filter = "PNG File|*.PNG|All Files|*.*";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            barstaitmLeftInfo.Caption = "Exporting";

            chartcCompareGPAs.ExportToImage(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmsubEAnamyGPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            barbtnitmAnalyseGPA_ItemClick(sender, e);

            saveFileDialog.Filter = "PNG File|*.PNG|All Files|*.*";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            barstaitmLeftInfo.Caption = "Exporting";

            chartcCompareGPAs.ExportToImage(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmsubESortmyGPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            barbtnitmSortGPA_ItemClick(sender, e);

            saveFileDialog.Filter = "PNG File|*.PNG|All Files|*.*";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            barstaitmLeftInfo.Caption = "Exporting";

            chartcCompareGPAs.ExportToImage(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png);

            barstaitmLeftInfo.Caption = "Ready";
        }

        private void barbtnitmCustomGPAMethods_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            xtraTabControl.SelectedTabPageIndex = 4;
        }

        private void barbtnitmSetMyName_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gridvCustomData.FocusedValue == null || gridvCustomData.FocusedColumn.AbsoluteIndex != 1)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show("Error, Please select the cell that contains your name in column 2.");
                return;
            }

            baredtitmMyName.EditValue = gridvCustomData.FocusedValue.ToString();
        }

        private void bsvbiNew_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmNew_ItemClick(sender, null);
        }

        private void bsvbiOpen_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmOpen_ItemClick(sender, null);
        }

        private void bsvbiSave_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmSave_ItemClick(sender, null);
        }

        private void bsvbiPrint_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmPrint_ItemClick(sender, null);
        }

        private void bsvbiExport_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            ribbonControl.SelectedPage = ribpagData;
        }

        private void bsvbiClose_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmNew_ItemClick(sender, null);
        }

        private void bsvbiSettings_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmSettings_ItemClick(sender, null);
        }

        private void bsvbiAbout_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            barbtnitmAbout_ItemClick(sender, null);
        }

        private void bsvbiExit_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            if (dtCustomData.Rows.Count>0)
            {
                DialogResult dr = DevExpress.XtraEditors.XtraMessageBox.Show("Do you want to save current data?", "Close", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (dr==DialogResult.Yes)
                {
                    barbtnitmSave_ItemClick(sender, null);
                }
                else if (dr==DialogResult.Cancel)
                {
                    return;
                }
            }
            this.Close();
        }
    }
}
