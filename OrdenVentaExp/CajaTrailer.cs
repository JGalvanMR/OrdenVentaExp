using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace OrdenVentaExp
{
    public partial class CajaTrailer : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd = new SqlCommand();
        SqlDataReader reader;

        public CajaTrailer()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void txtfolio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                try
                {
                    thisConnection.Open();
                    cmnd = thisConnection.CreateCommand();
                    cmnd.CommandText = "select placacaja from tb_mstr_pedidos_exp where pdn_folio = '" + txtfolio.Text + "'";
                    reader = cmnd.ExecuteReader();
                    while (reader.Read())
                    {
                        txtct.Text = reader.GetSqlString(0).ToString();
                    }
                    if (reader.HasRows == false)
                    {
                        MessageBox.Show("No existe el número de pedido", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        thisConnection.Close();
                        txtfolio.Text = string.Empty;
                    }
                    reader.Close();
                    thisConnection.Close();
                }
                catch (SqlException ex)
                {                    
                    thisConnection.Close();
                    MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtfolio.Text = string.Empty;
                    return;
                }
            }
        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PBbtnguardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("¿Desea guardar los cambios?", "AVISO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    thisConnection.Open();
                    cmnd = thisConnection.CreateCommand();
                    cmnd.CommandText = "update tb_mstr_pedidos_exp set placacaja ='" + txtct.Text + "' where pdn_folio = '" + txtfolio.Text + "'";
                    reader = cmnd.ExecuteReader();
                    reader.Close();
                    thisConnection.Close();
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", txtfolio.Text, "CAMBIO DE PLACAS DEL TRAILER " + txtct.Text + " DE LA ORD. VTA. EXP. " + txtfolio.Text, "SIPGAB");
                    Utilerias.Class1.Nombre_equipo = txtct.Text;
                    this.Close();
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                thisConnection.Close();
            }
        }
    }
}
