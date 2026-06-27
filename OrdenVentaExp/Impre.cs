using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Printing;

namespace OrdenVentaExp
{
    public partial class Impre : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlCommand cmnd = new SqlCommand();
        SqlDataReader reader;

        public Impre()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void Impre_Load(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "select nom_impre from tb_Cat_impresoras";
                reader = cmnd.ExecuteReader();
                while (reader.Read())
                {
                    CBselecimp.Items.Add(reader.GetSqlString(0).ToString());
                }
                reader.Close();

                cmnd = thisConnection.CreateCommand();
                cmnd.CommandText = "select nom_impre from tb_Cat_impresoras where pred_impre ='S'";
                reader = cmnd.ExecuteReader();
                while (reader.Read())
                {
                    txtimpdeter.Text = reader.GetSqlString(0).ToString();
                }
                reader.Close();
                thisConnection.Close();
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        int id;
        private void CBselecimp_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Utilerias.Class1.Nombre_impresora = CBselecimp.SelectedItem.ToString().Trim();

            thisConnection.Open();
            cmnd = thisConnection.CreateCommand();
            cmnd.CommandText = "select id_impre from tb_Cat_impresoras where nom_impre = '" + CBselecimp.SelectedItem.ToString() + "'";
            reader = cmnd.ExecuteReader();
            while (reader.Read())
            {
                id = reader.GetSqlInt32(0).Value;
            }
            thisConnection.Close();
        }

        private void PBbtnguardar_Click(object sender, EventArgs e)
        {
            try 
            {
                if (CBselecimp.SelectedIndex != -1)
                {
                    thisConnection.Open();
                    cmnd = thisConnection.CreateCommand();
                    cmnd.CommandText = "update tb_Cat_impresoras set pred_impre = 'S' where nom_impre ='" + CBselecimp.SelectedItem.ToString() + "' and id_impre = '" + id + "'";
                    reader = cmnd.ExecuteReader();
                    reader.Close();

                    cmnd = thisConnection.CreateCommand();
                    cmnd.CommandText = "UPDATE Tb_Cat_Impresoras SET pred_impre = 'N' WHERE  id_impre!='" + id + "'";
                    reader = cmnd.ExecuteReader();
                    reader.Close();
                    thisConnection.Close();
                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1.1", CBselecimp.SelectedItem.ToString(), "CAMBIO DE IMPRESORA EN ORD. VTA. EXP", "SIPGAB");
                    MessageBox.Show("DATOS GRABADOS", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtimpdeter.Text = CBselecimp.SelectedItem.ToString();
                }
                else
                {
                    MessageBox.Show("NO SE HA SELECCIONADO NINGUNA IMPRESORA", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (SqlException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void PBbtnsalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CBselecimp_KeyDown(object sender, KeyEventArgs e)
        {
            CBselecimp.DroppedDown = true;
        }
    }
}
