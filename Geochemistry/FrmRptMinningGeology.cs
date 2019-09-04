using Microsoft.Reporting.WinForms;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace Geochemistry
{
    public partial class FrmRptMinningGeology : Form
    {
        DataTable datosReporte = new DataTable();
        public FrmRptMinningGeology()
        {
            InitializeComponent();
        }

        public FrmRptMinningGeology(DataTable gridControlMuestreo)
        {
            InitializeComponent();
            datosReporte = gridControlMuestreo;
        }

        private void FrmRptDiarioMuestreo_Load(object sender, EventArgs e)
        {
            string reporte = Path.Combine(Application.StartupPath, @"Informes\RptControlMuestreo.rdlc");
            this.reportViewer1.LocalReport.ReportPath = reporte;
            this.reportViewer1.LocalReport.DataSources.Clear();
            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("ReporteControlMuestreoDataSet", datosReporte));
            this.reportViewer1.RefreshReport();
        }
    }
}
