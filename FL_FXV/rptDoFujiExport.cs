using System;
using System.Collections.Generic;
using DevExpress.Utils;
using DevExpress.XtraReports.UI;
using FL_FXV.Object;

namespace FL_FXV
{
    public partial class rptDoFujiExport : XtraReport
    {
        public rptDoFujiExport()
        {
            InitializeComponent();
        }
        public void SetData(DO_FUJI_XEROXDTO dtoMaster, List<DO_FUJI_XEROXDTO> lstChiTiet)
        {
            var lstData = new List<InPhieuDOFuji>();
            var data = new InPhieuDOFuji();
            data.MasterDTO = dtoMaster;
            data.LstChiTiet = lstChiTiet;
            lstData.Add(data);
            DataSource = lstData;
        }

    }
    public class InPhieuDOFuji
    {
        public InPhieuDOFuji()
        {
            LstChiTiet = new List<DO_FUJI_XEROXDTO>();
            MasterDTO = new DO_FUJI_XEROXDTO();
        }
        public List<DO_FUJI_XEROXDTO> LstChiTiet
        {
            get;
            set;
        }
        public DO_FUJI_XEROXDTO MasterDTO
        {
            get;
            set;
        }
    }
}
