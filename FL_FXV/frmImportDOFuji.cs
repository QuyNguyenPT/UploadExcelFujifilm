using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraReports.UI;
using FL_FXV.Helper;
using FL_FXV.Object;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FL_FXV
{
    public partial class frmImportDOFuji : Form
    {
        public frmImportDOFuji()
        {
            InitializeComponent();
        }

        private void btnChonFileUpload_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog
                {
                    Filter = @"Excel Files (2007) (.xlsx)|*.xlsx|Excel Files (2003)(.xls)|*.xls",
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {

                    txtFilePath.Text = ofd.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex);
            }
        }

        private void btnDongY_Click(object sender, EventArgs e)
        {
            try
            {
                ReadExcel();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }

        private void ReadExcel()
        {
            grdUpLoad.DataSource = null;
            grdUpLoad.RefreshDataSource();

            if (txtFilePath.Text.Trim() == "")
            {
                MessageBox.Show("Bạn cần chọn file cần upload");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtTenSheet.Text))
            {
                MessageBox.Show("Vui lòng nhập tên sheet.");
                txtTenSheet.Focus();
                return;
            }
            var lstUpLoad = new List<DO_FUJI_XEROXDTO>();


            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(txtFilePath.Text.Trim());
            var aaa = xlWorkbook.Worksheets[1].Name;
            bool bCoSheet = false;
            for (int i = 1; i <= xlWorkbook.Worksheets.Count; i++)
            {
                if (xlWorkbook.Worksheets[i].Name == txtTenSheet.Text.Trim())
                {
                    bCoSheet = true;
                }
            }
            if (!bCoSheet)
            {
                MessageBox.Show("Không tìm thấy sheet: " + txtTenSheet.Text.Trim());
                return;
            }
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[txtTenSheet.Text.Trim()];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;

            var dtoCungOder = new DO_FUJI_XEROXDTO();
            for (int i = 2; i <= rowCount; i++)
            {
                var dtoUpload = new DO_FUJI_XEROXDTO();
                int k = 1;
                dtoUpload.ODER_NUMBER = GetValue<string>(i, ref k, xlRange);
                if (!string.IsNullOrWhiteSpace(dtoUpload.ODER_NUMBER))
                {
                    dtoUpload.DELIVERY_NO = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.CUST_NO = GetValue<string>(i, ref k, xlRange);
                    if (string.IsNullOrWhiteSpace(dtoUpload.CUST_NO))
                    {
                        MessageBox.Show(string.Format("Dòng {0} Cus No không được phép để trống.", i));
                        break;
                    }
                    var dtDeliveryDate = xlRange.Cells[i, k].Value;
                    if (dtDeliveryDate == null || string.IsNullOrWhiteSpace(dtDeliveryDate.ToString()))
                    {
                        MessageBox.Show(string.Format("Dòng {0} Delivery Date không được phép để trống.", i));
                        break;
                    }
                    dtoUpload.DELIVERY_DATE = GetValue<DateTime>(i, ref k, xlRange);
                    dtoUpload.CUST_NAME = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_ADDRESS = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIPING_INSTRUCTION = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.MOVE_ODER_NUMBER = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_ADDR1 = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_ADDR2 = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_ADDR3 = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_ADDR4 = GetValue<string>(i, ref k, xlRange);
                    dtoUpload.SHIP_TO_CITY = GetValue<string>(i, ref k, xlRange);

                    ObjectHelper.inJect(dtoUpload, dtoCungOder);
                }
                else
                {
                    ObjectHelper.inJect(dtoCungOder, dtoUpload);

                    k = 14;
                }

                dtoUpload.LINE_NO = GetValue<string>(i, ref k, xlRange);
                dtoUpload.ITEM_CODE = GetValue<string>(i, ref k, xlRange);
                if (string.IsNullOrWhiteSpace(dtoUpload.ITEM_CODE))
                {
                    MessageBox.Show(string.Format("Dòng {0} Item Code không được phép để trống.", i));
                    break;
                }
                dtoUpload.DESCRIPTION = GetValue<string>(i, ref k, xlRange);
                dtoUpload.SERIAL_NUMBER = GetValue<string>(i, ref k, xlRange);
                dtoUpload.SERIAL_OF_MACHINE = GetValue<string>(i, ref k, xlRange);

                dtoUpload.ORDER_QUANTITY = GetValue<decimal>(i, ref k, xlRange);
                dtoUpload.SHIPED_QUANTITY = GetValue<decimal>(i, ref k, xlRange);
                dtoUpload.BACK_ORDERED_QUANTITY = GetValue<decimal>(i, ref k, xlRange);
                dtoUpload.SERIAL_BCCS_NO_1 = GetValue<string>(i, ref k, xlRange);
                dtoUpload.Select = true;
                lstUpLoad.Add(dtoUpload);
            }
            var lstUploadSort = lstUpLoad.OrderBy(x => x.ODER_NUMBER).ToList();
            grdUpLoad.DataSource = lstUploadSort;
            grdUpLoad.RefreshDataSource();

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Upload thành công.");
        }
        public T GetValue<T>(int rowIndex, ref int colIndex, Excel.Range xlRange)
        {
            var t = typeof(T);
            var value = xlRange.Cells[rowIndex, colIndex].Value;
            colIndex++;

            object stringEmpty = "";
            if (value == null)
            {
                return t == typeof(string) ? (T)stringEmpty : default(T);
            }

            if (String.IsNullOrWhiteSpace(value.ToString()))
            {
                return t == typeof(string) ? (T)stringEmpty : default(T);
            }

            if (t == typeof(decimal) || t == typeof(double))
            {
                return Convert.ChangeType(value, t);
            }
            return Convert.ChangeType(value, t);
        }
        private void btnSaoChepDuLieuMau_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                grvFormMau.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.True;
                grvFormMau.SelectAll();
                grvFormMau.CopyToClipboard();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnDong_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        public static XtraReport CombineReport(List<XtraReport> reports)
        {
            if (reports.Count > 0)
            {
                var firstReport = reports[0];
                firstReport.CreateDocument();
                for (int i = 1; i < reports.Count; i++)
                {
                    reports[i].CreateDocument();
                    firstReport.Pages.AddRange(reports[i].Pages);
                }

                return firstReport;
            }
            return null;
        }

        private void btnIn_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                ObjectHelper.CommitGridToDataSource(grdUpLoad);
                var lstChiTiet = grdUpLoad.DataSource as List<DO_FUJI_XEROXDTO>;
                if (lstChiTiet == null || lstChiTiet.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu cần in.");
                    return;
                }
                var lstSelect = lstChiTiet.Where(x => x.Select).ToList();
                if (lstSelect == null || lstSelect.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn dữ liệu cần in.");
                    return;
                }
                var reports = new List<XtraReport>();
                var lstOderCode = lstSelect.Select(x => x.ODER_NUMBER).Distinct().ToList();
                var lstOrderSort = lstOderCode.OrderBy(x => x).ToList();
                foreach (var code in lstOrderSort)
                {
                    var lstDetail = lstChiTiet.Where(x => x.ODER_NUMBER == code).ToList();
                    var rpt = new rptDoFujiExport();
                    rpt.SetData(lstDetail[0], lstDetail);
                    reports.Add(rpt);
                }
                var rptCombie = CombineReport(reports);
                rptCombie.PrintingSystem.ContinuousPageNumbering = false;
                rptCombie.ShowPreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void grvUpLoad_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                GridView grid = sender as GridView;
                if (grid == null)
                {
                    return;
                }

                if (!grid.IsDataRow(e.RowHandle))
                {
                    return;
                }

                if (e.Column == null || (e.Column.FieldName.ToUpper() != "SELECT" && e.Column.FieldName.ToUpper() != "SELECTED"))
                {
                    return;
                }

                var dto = grid.GetRow(e.RowHandle) as DO_FUJI_XEROXDTO;
                if (dto == null)
                {
                    return;
                }

                if (e.Column.FieldName.ToUpper() == "SELECT")
                {
                    dto.Select = !dto.Select;
                }
                else if (e.Column.FieldName.ToUpper() == "SELECTED")
                {
                    dto.Select = !dto.Select;
                }

                grid.RefreshData();
                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void grvUpLoad_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyData == Keys.Space)
                {
                    GridView grid = sender as GridView;

                    if (grid != null)
                    {
                        int[] handleList = grid.GetSelectedRows();
                        if (handleList.Length > 0)
                        {
                            GridColumn focusColumn = grid.FocusedColumn;
                            int focusRowHandle = grid.FocusedRowHandle;

                            foreach (int i in handleList)
                            {
                                if (!grid.IsDataRow(i))
                                {
                                    continue;
                                }

                                //Nếu column đang focus chính là Column chọn
                                if (focusColumn != null && (focusColumn.FieldName.ToUpper() == "SELECT" || focusColumn.FieldName.ToUpper() == "SELECTED"))
                                {
                                    if (focusRowHandle == i)
                                    {
                                        grid.CloseEditor();
                                        continue;
                                    }
                                }

                                var dto = grid.GetRow(i) as DO_FUJI_XEROXDTO;
                                if (dto != null)
                                {
                                    dto.Select = !dto.Select;
                                }

                            }
                            grid.RefreshData();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnClear_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                var lstChiTiet = grdUpLoad.DataSource as List<DO_FUJI_XEROXDTO>;
                if (lstChiTiet == null || lstChiTiet.Count == 0)
                {
                    return;
                }
                foreach (var item in lstChiTiet)
                {
                    item.Select = false;
                }
                grdUpLoad.RefreshDataSource();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnChonTatCa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                var lstChiTiet = grdUpLoad.DataSource as List<DO_FUJI_XEROXDTO>;
                if (lstChiTiet == null || lstChiTiet.Count == 0)
                {
                    return;
                }
                foreach (var item in lstChiTiet)
                {
                    item.Select = true;
                }
                grdUpLoad.RefreshDataSource();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
