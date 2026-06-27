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
    public partial class Alta_Subclientes : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd1;
        SqlDataReader reader1;
        string subcnte = "";

        public Alta_Subclientes()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if ((txtcnte.Text.Trim().Length > 0) && (txtsubcnte.Text.Trim().Length > 0) && (txtnomsubcnte.Text.Trim().Length > 0))
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select entr_clave from tb_cat_entregas where cnte_clave = '" + txtcnte.Text + "' and entr_clave = '" + txtsubcnte.Text + "'";
                subcnte = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                if (subcnte.Trim() != "")
                {
                    thisConnection.Close();
                    MessageBox.Show("Ya esta dado de alta este subcliente: " + txtsubcnte.Text, "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "insert into tb_cat_entregas (cnte_clave, entr_clave, entr_nombre, entr_sufijo, entr_rfc, entr_telefono, entr_mail, entr_direccion, entr_colonia, "+
                                        "entr_ciudad, entr_estado, entr_cp, entr_contacto, entr_cvesuc, entr_cedis) values " +
                                        "('" + txtcnte.Text + "', '" + txtsubcnte.Text + "', '" + txtnomsubcnte.Text + "', '', '', '', '', '', '', '', '', '', '', '', '')";
                    reader1 = cmnd1.ExecuteReader();
                    thisConnection.Close();
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", txtcnte.Text, "NUEVO SUBCLIENTE " + txtsubcnte.Text + " " + txtnomsubcnte.Text, "SIPGAB");
                    MessageBox.Show("Subcliente dado de alta", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtcnte.Clear(); txtnomsubcnte.Clear(); txtsubcnte.Clear();
                }
            }
            else
                MessageBox.Show("No se puede guardar la información ya que alguno de los campos esta vacio", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
