using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace OrdenVentaExp
{
    public partial class FrmCompleExp : Form
    {
        SqlConnection thisconnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlConnection thisConnectionUERP = new SqlConnection(Utilerias.Class1.ConnectionStringUERP);
        public DataTable Info = new DataTable();
        DataTable CatUERP = new DataTable();
        public string Destino = "", Domicilio = "", Folio = "";

        public FrmCompleExp()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void FrmCompleExp_Load(object sender, EventArgs e)
        {
            DgDatos.DataSource = Info;
            if (!Info.Columns.Contains("FraccionIni"))
            {
                Info.Columns.Add("FraccionIni", typeof(string));
                Info.Columns.Add("Peso2", typeof(string));
            }
            foreach (DataRow row in Info.Rows)
            {
                row["FraccionIni"] = row["Fraccion"];
                row["Peso2"] = row["Peso"];
            }
            LblDes.Text = Destino;
            LblDire.Text = Domicilio;
            //07/10/2022
            thisConnectionUERP.Open();
            string Cadena = "select a.artcod, a.artnom, a.artcantuni, a.unisimaux, b.CatUnidadI, c.catprodsid " +
                            "from articulos a left join saterpunid b on b.satunisim = a.unisimb " +
                            "left join saterparti c on a.artcod = c.SatArtCod and c.satartcod = a.artcod and a.empcod = c.satempcod " +
                            "where a.empcod = '60'  order by a.artcod";
            SqlDataAdapter da = new SqlDataAdapter(Cadena, thisConnectionUERP);
            DataSet ds = new DataSet();
            da.Fill(ds, "CatUERP");
            CatUERP = ds.Tables["CatUERP"];
            thisConnectionUERP.Close();
            //fin 07/10/2022
            lbtil.Text = lbtil.Text + "  F- " + Folio;
            Formato();
        }

        private void Formato()
        {
            //Int32 i = 0; 
            //foreach (DataGridViewRow row in DgDatos.Rows)
            //{
            //    decimal cant = Convert.ToDecimal(DgDatos.Rows[i].Cells["Importe"].Value) - Convert.ToDecimal(DgDatos.Rows[i].Cells["Total"].Value);
            //    decimal cant1 = Math.Abs(cant);
            //    if (cant1 > 10) 
            //    {
            //        DgDatos.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
            //    }
            //    i++;
            //}

            DgDatos.Columns[1].HeaderText = "Producto";
            DgDatos.Columns[2].HeaderText = "Fraccion";
            DgDatos.Columns[3].HeaderText = "Aduana";
            DgDatos.Columns[4].HeaderText = "Cajas";
            DgDatos.Columns[5].HeaderText = "Peso";
            DgDatos.Columns[6].HeaderText = "Kilos";
            DgDatos.Columns[7].HeaderText = "Precio Uni";
            DgDatos.Columns[8].HeaderText = "Importe";
            DgDatos.Columns[9].HeaderText = "Kilos Originales";
            //DgDatos.Columns[0].Visible = false;
            //DgDatos.Columns[9].Visible = false;
            DgDatos.Columns[1].Width = 350;
            DgDatos.Columns[2].Width = 70;
            DgDatos.Columns[3].Width = 50;
            DgDatos.Columns[4].Width = 60;
            DgDatos.Columns[5].Width = 60;
            DgDatos.Columns[6].Width = 80;
            DgDatos.Columns[7].Width = 80;
            DgDatos.Columns[8].Width = 100;
            DgDatos.Columns[9].Width = 100;
            DgDatos.Columns[4].DefaultCellStyle.Format = "#,##0";
            DgDatos.Columns[5].DefaultCellStyle.Format = "#,##0.###0";
            DgDatos.Columns[6].DefaultCellStyle.Format = "#,##0.###0";
            DgDatos.Columns[7].DefaultCellStyle.Format = "#,##0.###0";
            DgDatos.Columns[8].DefaultCellStyle.Format = "#,##0.###0";
            DgDatos.Columns[9].DefaultCellStyle.Format = "#,##0.###0";
            DgDatos.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DgDatos.Columns[1].ReadOnly = true;
            //DgDatos.Columns[2].ReadOnly = true;
            //DgDatos.Columns[3].ReadOnly = true;
            DgDatos.Columns[4].ReadOnly = true;
            //DgDatos.Columns[5].ReadOnly = true;
            DgDatos.Columns[6].ReadOnly = true;
            DgDatos.Columns[8].ReadOnly = true;
            DgDatos.Columns[10].Visible = false;
            DgDatos.Columns[11].Visible = false;

        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PbxExcel_Click(object sender, EventArgs e)
        {
            Excel.Range r;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Application.Workbooks.Add();

            //Imagen a agregar (notese el tamaño 500×450 configurable)
            //r = excel.Range[excel.Cells[1, 1], excel.Cells[7, 1]];
            //r.Worksheet.Shapes.AddPicture(@"C:\SisGabWeb\LogoSunset.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 1f, 1f, 100, 70);
            //r = excel.Range[excel.Cells[1, 1], excel.Cells[7, 1]];
            //r = excel.Range[excel.Columns[1], excel.Columns[1]];
            //r.RowHeight = 50;
            //int count = MiDataGrid.ColumnCount;
            //excel.Range[excel.Cells[1, 1], excel.Cells[1, 7]].Merge();
            //excel.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            excel.Cells[1, 1] = "CVE (SIPGAB)";
            excel.Cells[1, 2] = "CODIGO SAT (UERP)";
            excel.Cells[1, 3] = "CLAVE UNIDAD";
            excel.Cells[1, 4] = "CLAVE PESO";
            excel.Cells[1, 5] = "PRODUCTO";
            excel.Cells[1, 6] = "TIPO DE MATERIA";
            excel.Cells[1, 7] = "REGULADA POR COFEPRIS";
            excel.Cells[1, 8] = "TIPO DE EXPORTACION";
            excel.Cells[1, 9] = "FRACCION";
            excel.Cells[1, 10] = "ADUANA";
            excel.Cells[1, 11] = "CAJAS";
            excel.Cells[1, 12] = "PESO";
            excel.Cells[1, 13] = "KILOS";
            excel.Cells[1, 14] = "PRECIO UNI";
            excel.Cells[1, 15] = "IMPORTE";
            excel.Cells[1, 16] = "KILOS ORIGINALES";
            excel.Cells[1, 17] = "DESTINATARIO";
            excel.Cells[1, 18] = "DIRECCIÓN";
            r = excel.Range[excel.Cells[1, 1], excel.Cells[1, 18]];
            r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            //r = excel.Range[excel.Cells[7, 1], excel.Cells[13, 4]];
            //r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //r = excelWorksheet.get_Range("A1", "R1");
            //r.Font.Bold = true;



            //r = excel.Range[excel.Cells[7, 1], excel.Cells[7, 11]];
            //r.Font.Bold = true;
            //for (i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //    dgc = dataGridView1.Rows[i].Cells[0];
            //    celda = ((String)dgc.Value) + "\r\n";
            //    textBox1.Text += celda.Replace(".", ",");
            //}

            //agrega las filas a excel
            int i = 2;
            int j = 0;
            int TotCjs = 0;
            foreach (DataGridViewRow row in DgDatos.Rows)
            {
                //No.	Fecha	Turno	Sup. Carga	Temp	Hr. Llegó	Hr. Entro	Hr. Salio	Tiempo Total	Chofer	Destino	Ini. Carga	Fin. Carga	Tiempo Carga	Anden	Transporte	Placa Caja	Placa Trailer	Radio
                excel.Cells[i, 1] = DgDatos.Rows[j].Cells[0].Value.ToString(); // Cve
                foreach (DataRow Row in CatUERP.Select("ArtCod = '" + DgDatos.Rows[j].Cells[0].Value.ToString() + "'"))
                {
                    excel.Cells[i, 2] = Row["catprodsid"].ToString(); // Cve SAT
                    excel.Cells[i, 3] = Row["CatUnidadI"].ToString(); // Unidad SAt
                }
                excel.Cells[i, 4] = "KGS"; // Cve Peso
                excel.Cells[i, 5] = DgDatos.Rows[j].Cells[1].Value.ToString(); // 
                excel.Cells[i, 6] = "03-Materia terminada";
                excel.Cells[i, 7] = "NO";
                excel.Cells[i, 8] = "DEFINITIVA";
                excel.Cells[i, 9] = "'" + DgDatos.Rows[j].Cells[2].Value.ToString(); // 
                excel.Cells[i, 10] = DgDatos.Rows[j].Cells[3].Value.ToString(); // 
                excel.Cells[i, 11] = Convert.ToDecimal(DgDatos.Rows[j].Cells[4].Value).ToString("##,###"); // 
                excel.Cells[i, 12] = Convert.ToDecimal(DgDatos.Rows[j].Cells[5].Value).ToString("##,###.00"); // 
                excel.Cells[i, 13] = Convert.ToDecimal(DgDatos.Rows[j].Cells[6].Value).ToString("##,###.00"); // 
                excel.Cells[i, 14] = Convert.ToDecimal(DgDatos.Rows[j].Cells[7].Value).ToString("##,###.00"); // 
                excel.Cells[i, 15] = Convert.ToDecimal(DgDatos.Rows[j].Cells[8].Value).ToString("##,###.00"); // 
                excel.Cells[i, 16] = Convert.ToDecimal(DgDatos.Rows[j].Cells[9].Value).ToString("##,###.00"); // 
                //excel.Cells[i, 11] = DgDatos.Rows[j].Cells[10].Value.ToString(); // 
                excel.Cells[i, 17] = LblDes.Text;
                excel.Cells[i, 18] = LblDire.Text;
                i++;
                j++;

            }
            //r = excel.Range[excel.Cells[9, 1], excel.Cells[i - 1, 4]];
            //r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //r = excel.Range[excel.Cells[i, 2], excel.Cells[i, 3]];
            //r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            excel.Columns.AutoFit();
            //excel.Rows.AutoFit();
            MessageBox.Show("Archivo Generado Con Exito!!!!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            excel.Visible = true;

        }


    }
}
