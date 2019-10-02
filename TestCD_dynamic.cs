using DevExpress.XtraGrid.Views.BandedGrid;
using FozzySystems;
using FozzySystems.Proxy;
using FozzySystems.Types.Contracts;
using FozzySystems.Utils;
using SuppliersPortal;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace ApplicationRegistry.SuppliersPortal.Reports
{
    public partial class TestCD_dynamic : Form
    {
        private DataTable dtImport;

        public TestCD_dynamic()
        {
            InitializeComponent();
            GetBusinessList();

            barBtnSave.Enabled = false;

            advBandedGridView1.ShownEditor += gridView1_ShownEditor;

            dtImport = new DataTable("Import");
            dtImport.Columns.Add("errorCode", typeof(string));
            
        }

        private void gridView1_ShownEditor(object sender, EventArgs e)
        {
            if (advBandedGridView1.FocusedColumn.AbsoluteIndex > 3)
            {
                advBandedGridView1.ActiveEditor.MouseUp += new MouseEventHandler(ActiveEditor_MouseUp);
            }
        }

        private void ActiveEditor_MouseUp(object sender, MouseEventArgs e)
        {
            var myCell = advBandedGridView1.GetFocusedValue();

            using (Login form = new Login())
            {
                if (myCell != DBNull.Value)
                {
                    form.GetNotEmptyCell((string)myCell);
                }

                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    string columnName = advBandedGridView1.FocusedColumn.FieldName;

                    advBandedGridView1.SetRowCellValue(advBandedGridView1.FocusedRowHandle, columnName, form.myValue);

                    barBtnSave.Enabled = true;
                }
            }
        }

        private void GetBusinessList()
        {
            FZCoreProxy.ExecuteAsync(this, GotCallback, "SuppliersPortal.Reports.TestCD@filters", null);
        }

        private void GotCallback(IDefaultContract o, object uo)
        {
            try
            {
                if (o != null && o.errorCode != ErrorCodes.OK)
                {
                    return;
                }

                DataTable dataBusiness = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[0];
                DataTable dataCG = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[1];
                DataTable dataRegion = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[2];
                DataTable dataRespons = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[3];

                if (dataBusiness.Rows.Count > 0)
                {
                    ccbBusiness.Properties.DataSource = dataBusiness;
                    ccbBusiness.Properties.DisplayMember = "businessName";
                    ccbBusiness.Properties.ValueMember = "businessId";
                }
                else
                    ccbBusiness.Properties.DataSource = null;

                if (dataCG.Rows.Count > 0)
                {
                    ccbCommodityGroups.Properties.DataSource = dataCG;
                    ccbCommodityGroups.Properties.DisplayMember = "commodityGroupName";
                    ccbCommodityGroups.Properties.ValueMember = "commodityGroupId";
                }
                else
                    ccbCommodityGroups.Properties.DataSource = null;

                if (dataRegion.Rows.Count > 0)
                {
                    ccbMacroRegion.Properties.DataSource = dataRegion;
                    ccbMacroRegion.Properties.DisplayMember = "macroRegionNameRu";
                    ccbMacroRegion.Properties.ValueMember = "macroRegionId";
                }
                else
                    ccbMacroRegion.Properties.DataSource = null;

                if (dataRespons.Rows.Count > 0)
                {
                    ccbRespons.Properties.DataSource = dataRespons;
                    ccbRespons.Properties.DisplayMember = "responsibleName";
                    ccbRespons.Properties.ValueMember = "responsibleName";
                }
                else
                    ccbMacroRegion.Properties.DataSource = null;

                CheckAllItem();

            }
            catch (Exception ex)
            {
                MB.error(ex);
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            CheckAllItem();
            btnApply.PerformClick();
        }

        private void CheckAllItem()
        {
            ccbBusiness.CheckAll();
            ccbCommodityGroups.CheckAll();
            ccbMacroRegion.CheckAll();
            ccbRespons.CheckAll();

            ccbBusiness.RefreshEditValue();
            ccbCommodityGroups.RefreshEditValue();
            ccbMacroRegion.RefreshEditValue();
            ccbRespons.RefreshEditValue();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            FZCoreProxy.ExecuteAsync(this, LoadDataCallback, "SuppliersPortal.Reports.TestCD", MyFilters(), null);
        }

        private void LoadDataCallback(IDefaultContract o, object uo)
        {
            advBandedGridView1.Columns.Clear();
            advBandedGridView1.OptionsView.ColumnAutoWidth = false;

            try
            {
                if (o != null && o.errorCode != ErrorCodes.OK)
                {
                    return;
                }
                
                DataTable dataBands = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[0];
                DataTable dataColumns = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[1];
                DataTable dataGrid = FZCoreProxy.ExecuteDataSet(o as ExecuteContract).Tables[2];

                AddBandsAndCoulumns(dataBands, dataColumns);

                if (dataGrid.Rows.Count > 0)
                    reportFormGrid1.DataSource = dataGrid;
                else
                    reportFormGrid1.DataSource = null;

                advBandedGridView1.OptionsView.ColumnAutoWidth = true;
            }
            catch (Exception ex)
            {
                MB.error(ex);
            }
        }

        private void AddBandsAndCoulumns(DataTable dataBands, DataTable dataColumns)
        {
            advBandedGridView1.Bands.Clear();
            var list = new List<KeyValuePair<int?, GridBand>>();
            var listBand = new List<KeyValuePair<int, GridBand>>();

            foreach (DataRow row in dataBands.Rows)
            {
                GridBand band = new GridBand();

                if (String.IsNullOrEmpty(row[2].ToString()))
                {
                    band.Caption = row[1].ToString();
                    band.Name = band.Caption;
                    advBandedGridView1.Bands.Add(band);

                    list.Add(new KeyValuePair<int?, GridBand>(null, band));
                    listBand.Add(new KeyValuePair<int, GridBand>((int)row[0].ToInt(), band));
                }
                else
                {
                    band.Caption = row[1].ToString();
                    band.Name = band.Caption;

                    list.Add(new KeyValuePair<int?, GridBand>((int)row[2].ToInt(), band));
                    listBand.Add(new KeyValuePair<int, GridBand>((int)row[0].ToInt(), band));
                }
            }

            foreach (var value in list)
            {
                if (value.Key != null)
                {
                    foreach (var valueBand in listBand)
                    {
                        if (value.Key == valueBand.Key)
                        {
                            valueBand.Value.Children.Add(value.Value);
                        }

                        if (valueBand.Value == value.Value)
                        {
                            foreach (DataRow rowCol in dataColumns.Rows)
                            {
                                if (valueBand.Key == rowCol[3].ToInt())
                                {
                                    AddColumn(value.Value, rowCol[0].ToString(), rowCol[1].ToString());
                                }
                            }
                        }
                    }
                }
            }
        }

        private void AddColumn(GridBand band, string columnName, string columnCaption)
        {
            BandedGridColumn bandedGridColumn = new BandedGridColumn();
            bandedGridColumn.FieldName = columnName;
            bandedGridColumn.Caption = columnCaption;
            advBandedGridView1.Columns.Add(bandedGridColumn);
            bandedGridColumn.Visible = true;
            bandedGridColumn.OwnerBand = band;
        }

        private string MyFilters()
        {
            var valuesB = ccbBusiness.Properties.Items.GetCheckedValues();
            var valuesCG = ccbCommodityGroups.Properties.Items.GetCheckedValues();
            var valuesMR = ccbMacroRegion.Properties.Items.GetCheckedValues();
            var valuesRes = ccbRespons.Properties.Items.GetCheckedValues();
            string myFlag = "4";
            string resultB = String.Join(",", valuesB.ToArray());
            string resultCG = String.Join(",", valuesCG.ToArray());
            string resultMR = String.Join(",", valuesMR.ToArray());
            string resultRes = String.Join(",", valuesRes.ToArray());

            DataTable table = new DataTable("Filters");

            table.Columns.Add("myFlag", typeof(string));
            table.Columns.Add("businessesId", typeof(string));
            table.Columns.Add("commodityGroupsId", typeof(string));
            table.Columns.Add("macroRegionName", typeof(string));
            table.Columns.Add("responsibleName", typeof(string));

            DataRow dataRow = table.NewRow();

            dataRow["myFlag"] = myFlag;
            dataRow["businessesId"] = resultB;
            dataRow["commodityGroupsId"] = resultCG;
            dataRow["macroRegionName"] = resultMR;
            dataRow["responsibleName"] = resultRes;

            table.Rows.Add(dataRow);

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);
            dataSet.Namespace = "";

            return Serialization.Serialize(dataSet);
        }

        private void barBtnSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SaveDataFromGrid();
        }

        private void SaveDataFromGrid()
        {
            advBandedGridView1.CloseEditor();

            FZCoreProxy.ExecuteAsync(this, (contract, state) =>
            {
                var c = contract as IDefaultContract;
                if (c != null && c.errorCode != ErrorCodes.OK)
                    MB.error(new Exception(c.errorString).ToString());
                DataSet dataSet = FZCoreProxy.ExecuteDataSet(c as ExecuteContract);
            }, "SuppliersPortal.Reports.TestCD@Write", Serialization.Serialize(reportFormGrid1.DataSource), null);

            barBtnSave.Enabled = false;
        }

        private void barBtnSaveToExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (reportFormGrid1.DataSource != null)
            {
                advBandedGridView1.ExportToExcelFromGrid("TestCD");
            }
        }

        private void barBtnLoadFromExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            using (OpenFileDialog ofdImport = new OpenFileDialog()
            {
                Filter = "|*.xls"
            })
            {
                if (ofdImport.ShowDialog() == DialogResult.OK)
                {
                    string select = "SELECT * FROM [Портал постачальників$A2:M2001] ";
                

                    string ConnectionStr = ImportTools.GetConnectionString(ofdImport.FileName);

                    using (OleDbConnection cn = new OleDbConnection(ConnectionStr))
                    {
                        using (OleDbDataAdapter daAcess = new OleDbDataAdapter(select, ConnectionStr))
                        {
                            try
                            {
                                cn.Open();
                                if (dtImport.Rows.Count != 0)
                                    dtImport.Rows.Clear();

                                daAcess.Fill(dtImport);
                            }
                            catch (Exception ex)
                            {
                                MB.error(ex.Message);
                            }
                        }
                    }

                    if (dtImport == null || dtImport.Rows.Count == 0)
                    {
                        MB.error("Немає даних для імпорту.");
                        return;
                    }

                    dtImport.AcceptChanges();
                    SaveDataFromExcel();
                }
            }
        }

        private void SaveDataFromExcel()
        {
            using (UserDataExcel userDataExcel = new UserDataExcel())
            {
                FZCoreProxy.ExecuteAsync(this, (contract, state) =>
                {
                    var c = contract as IDefaultContract;
                    if (c != null && c.errorCode != ErrorCodes.OK)
                        MB.error(new Exception(c.errorString).ToString());
                    DataTable data = new DataTable();
                    data = FZCoreProxy.ExecuteDataSet(c as ExecuteContract).Tables[0];
                    userDataExcel.GetDataTable(data);
                }, "SuppliersPortal.Reports.TestCD@Edit", Serialization.Serialize(dtImport), null);

                userDataExcel.ShowDialog(this);
            }
        }

        private void barBtnUpdateHistory_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            using (UpdateHistory form = new UpdateHistory())
            {
                form.ShowDialog(this);
            }
        }
    }
}
