using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OrdenVentaExp
{
    public partial class venta_especial : Form
    {
        public static string producto = "", descripcion = "";

        public venta_especial()
        {
            InitializeComponent();
            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        private void venta_especial_Load(object sender, EventArgs e)
        {
            txtclaveprod.Text = producto;
            txtdescripcion.Text = descripcion;
            OrdVenExp.especial_venta = txtdescripcion.Text;
            txtdescripcion.Select(descripcion.Length,0);
            txtdescripcion.Focus();
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bntsave_Click(object sender, EventArgs e)
        {
            OrdVenExp.especial_venta = txtdescripcion.Text;
            this.Close();
        }

        private void txtdescripcion_TextChanged(object sender, EventArgs e)
        {
            lbnumchar.Text = "Quedan " + (Convert.ToInt32(txtdescripcion.MaxLength) - Convert.ToInt32(txtdescripcion.Text.Length)) + " caracteres por escribir";
        }

        private void txtdescripcion_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
            }
        }
    }
}
