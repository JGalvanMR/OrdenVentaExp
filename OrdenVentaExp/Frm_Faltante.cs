using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;


namespace OrdenVentaExp
{
    public partial class Frm_Faltante : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd3;

        public DataTable datos = new DataTable();
        public string mped = "", mcnte = "", mnomcte = "", mfec = "", muser = "";

        public Frm_Faltante()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            Form.ActiveForm.Close();
        }

        private void Frm_Faltante_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < datos.Rows.Count; i++)
            {
                DGDatos.Rows.Add(datos.Rows[i][0].ToString(), datos.Rows[i][1].ToString(), datos.Rows[i][2].ToString(), datos.Rows[i][3].ToString());
            }

            thisConnection.Open();

            string Cadena = "SELECT * FROM TB_CAT_FALTANTE ORDER BY CAT_MOTIVO";
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnection;
            SqlDataReader Cata;
            Cata = cmd.ExecuteReader();

            DataGridViewComboBoxColumn dgvCmb1 = new DataGridViewComboBoxColumn();
            DataGridViewComboBoxColumn dgvCmb2 = new DataGridViewComboBoxColumn();
            DataGridViewComboBoxColumn dgvCmb3 = new DataGridViewComboBoxColumn();
            dgvCmb1.HeaderText = "Motivo";
            dgvCmb1.Name = "Motivo";
            dgvCmb2.HeaderText = "Area";
            dgvCmb2.Name = "Area";
            dgvCmb3.HeaderText = "Responsable";
            dgvCmb3.Name = "Responsable";
            while (Cata.Read())
            {
                if (Cata["cat_motivo"].ToString().Trim().Length > 0)
                    dgvCmb1.Items.Add(Cata["cat_motivo"].ToString());
                if (Cata["cat_area"].ToString().Trim().Length > 0)
                    dgvCmb2.Items.Add(Cata["cat_area"].ToString());
                if (Cata["cat_responsable"].ToString().Trim().Length > 0)
                    dgvCmb3.Items.Add(Cata["cat_responsable"].ToString());
            }
            DGDatos.Columns.Add(dgvCmb1);
            DGDatos.Columns.Add(dgvCmb2);
            DGDatos.Columns.Add(dgvCmb3);
            DGDatos.Columns[4].Width = 300;
            DGDatos.Columns[5].Width = 130;
            DGDatos.Columns[6].Width = 150;

            LblCnte.Text = Program.MyGlobal.PubCnte;
            LblPedido.Text = Program.MyGlobal.PubPed;
            LblNomCnte.Text = Program.MyGlobal.PubNomCnte;
            LblFecha.Text = Program.MyGlobal.PubFec;
            LblAnio.Text = Program.MyGlobal.PubFec.ToString().Substring(6, 4);
            DateTime Fpro = Convert.ToDateTime(Program.MyGlobal.PubFec);
            Cadena = "Select * FROM TB_CAT_SEMANAS WHERE '" + Fpro.ToString("dd/MM/yyyy") + "' >= FECHA1 AND '" + Fpro.ToString("dd/MM/yyyy") + "' <= FECHA2 ";
            //MessageBox.Show(Cadena);
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnection;
            SqlDataReader sem;
            sem = cmd.ExecuteReader();
            while (sem.Read())
            {
                LblSemana.Text = sem["semana"].ToString();
            }
            thisConnection.Close();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            DGDatos.CurrentCell = DGDatos.Rows[DGDatos.Rows.Count - 1].Cells["Clave"];
            foreach (DataGridViewRow Rows in DGDatos.Rows)
            {
                if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[4].FormattedValue.ToString()))
                {
                    MessageBox.Show("El Campo Motivo no se puede quedar Vacio del Producto" + System.Environment.NewLine + DGDatos.CurrentRow.Cells[1].Value.ToString());
                    return;
                }
                //else
                //    MessageBox.Show(DGDatos.CurrentRow.Cells[4].Value.ToString());
                if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[5].FormattedValue.ToString()))
                {
                    MessageBox.Show("El Campo Area no se puede quedar Vacio del Producto" + System.Environment.NewLine + DGDatos.CurrentRow.Cells[1].Value.ToString());
                    return;
                }
                //else
                //    MessageBox.Show(DGDatos.CurrentRow.Cells[5].Value.ToString());
                if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[6].FormattedValue.ToString()))
                {
                    MessageBox.Show("El Campo Responsable no se puede quedar Vacio del Producto" + System.Environment.NewLine + DGDatos.CurrentRow.Cells[1].Value.ToString());
                    return;
                }
                //else
                //    MessageBox.Show(DGDatos.CurrentRow.Cells[6].Value.ToString());

            }
            thisConnection.Open();
            //cmnd3 = thisConnection.CreateCommand();
            string mSql = "";
            try
            {
                int i = 0;
                foreach (DataGridViewRow iRows in DGDatos.Rows)
                {
                    string mnomprod = Convert.ToString(DGDatos.CurrentRow.Cells["Nombre"].Value).Trim();
                    mnomprod = mnomprod.Replace("'", " ");
                    string cadena = "insert into tb_faltante_prod (fecha,anio,semana,cnte_clave,prod_clave,cnte_nombre,prod_nombre,cajas,importe,pdn_folio,pdn_tipo,fcn_folio,motivo,area,resp_area, usuario) " +
                                        "values  ('" + LblFecha.Text + "','" + LblAnio.Text + "','" + LblSemana.Text + "','" + LblCnte.Text + "','" + Convert.ToString(DGDatos.Rows[i].Cells["Clave"].Value) + "','" +
                                        LblNomCnte.Text + "','" + mnomprod + "','" + Convert.ToDecimal(DGDatos.Rows[i].Cells["Cajas"].Value) + "','" + Convert.ToDecimal(DGDatos.Rows[i].Cells["importe"].Value) + "','" +
                                        LblPedido.Text + "','EXP','0','" + Convert.ToString(DGDatos.Rows[i].Cells[4].Value).Trim() + "','" + Convert.ToString(DGDatos.Rows[i].Cells[5].Value).Trim() + "','" +
                                        Convert.ToString(DGDatos.Rows[i].Cells[6].Value).Trim() + "','" + muser + "')";
                    SqlCommand cmdfal = new SqlCommand(cadena, thisConnection);
                    mSql = cadena.Replace("'", " ");
                    cmdfal.ExecuteNonQuery();
                    i++;
                }
                thisConnection.Close();
                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", LblPedido.Text, "FALTANTE DE PRODUCTO EN LA ORD VTA NAL " + LblPedido.Text, "SIPGAB");
                MessageBox.Show("Datos Gabados", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (SqlException ex)
            {
                 
                thisConnection.Close();
                Utilerias.Class1.SendMail("ricardo.cortes@mrlucky.com.mx", "ricardo.cortes", "rcedillo", ex.ToString().Trim() + " Pedido: " + LblPedido.Text + " User: " + muser + " Consulta:  " + mSql);
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void DGDatos_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "El valor de DataGridViewComboBoxCell no es válido." || e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = DGDatos.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)DGDatos.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)DGDatos.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
                }
            }
        }

        private void DGDatos_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 2)
                return;
            if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[4].FormattedValue.ToString()))
                MessageBox.Show("El Campo Motivo no se puede quedar Vacio");
            else
                MessageBox.Show(DGDatos.CurrentRow.Cells[4].Value.ToString());
            if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[5].FormattedValue.ToString()))
                MessageBox.Show("El Campo Area no se puede quedar Vacio");
            else
                MessageBox.Show(DGDatos.CurrentRow.Cells[5].Value.ToString());
            if (string.IsNullOrEmpty(DGDatos.CurrentRow.Cells[6].FormattedValue.ToString()))
                MessageBox.Show("El Campo Responsable no se puede quedar Vacio");
            else
                MessageBox.Show(DGDatos.CurrentRow.Cells[6].Value.ToString());
        }
    }
}
