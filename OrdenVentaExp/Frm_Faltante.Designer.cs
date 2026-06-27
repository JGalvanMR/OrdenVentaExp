namespace OrdenVentaExp
{
    partial class Frm_Faltante
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_Faltante));
            this.Cajas = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LblAnio = new System.Windows.Forms.Label();
            this.LblSemana = new System.Windows.Forms.Label();
            this.LblNomCnte = new System.Windows.Forms.Label();
            this.LblCnte = new System.Windows.Forms.Label();
            this.Clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LblPedido = new System.Windows.Forms.Label();
            this.Importe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LblFecha = new System.Windows.Forms.Label();
            this.DGDatos = new System.Windows.Forms.DataGridView();
            this.PBbtnsalir = new System.Windows.Forms.PictureBox();
            this.BtnSave = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.DGDatos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BtnSave)).BeginInit();
            this.SuspendLayout();
            // 
            // Cajas
            // 
            this.Cajas.HeaderText = "Cajas";
            this.Cajas.Name = "Cajas";
            this.Cajas.ReadOnly = true;
            this.Cajas.Width = 50;
            // 
            // Nombre
            // 
            this.Nombre.Frozen = true;
            this.Nombre.HeaderText = "Nombre";
            this.Nombre.Name = "Nombre";
            this.Nombre.ReadOnly = true;
            this.Nombre.Width = 300;
            // 
            // LblAnio
            // 
            this.LblAnio.AutoSize = true;
            this.LblAnio.BackColor = System.Drawing.Color.Transparent;
            this.LblAnio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblAnio.ForeColor = System.Drawing.Color.White;
            this.LblAnio.Location = new System.Drawing.Point(658, 33);
            this.LblAnio.Name = "LblAnio";
            this.LblAnio.Size = new System.Drawing.Size(23, 13);
            this.LblAnio.TabIndex = 102;
            this.LblAnio.Text = "....";
            // 
            // LblSemana
            // 
            this.LblSemana.AutoSize = true;
            this.LblSemana.BackColor = System.Drawing.Color.Transparent;
            this.LblSemana.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblSemana.ForeColor = System.Drawing.Color.White;
            this.LblSemana.Location = new System.Drawing.Point(658, 16);
            this.LblSemana.Name = "LblSemana";
            this.LblSemana.Size = new System.Drawing.Size(23, 13);
            this.LblSemana.TabIndex = 101;
            this.LblSemana.Text = "....";
            // 
            // LblNomCnte
            // 
            this.LblNomCnte.AutoSize = true;
            this.LblNomCnte.BackColor = System.Drawing.Color.Transparent;
            this.LblNomCnte.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNomCnte.ForeColor = System.Drawing.Color.White;
            this.LblNomCnte.Location = new System.Drawing.Point(71, 20);
            this.LblNomCnte.Name = "LblNomCnte";
            this.LblNomCnte.Size = new System.Drawing.Size(71, 20);
            this.LblNomCnte.TabIndex = 100;
            this.LblNomCnte.Text = "Nombre";
            // 
            // LblCnte
            // 
            this.LblCnte.AutoSize = true;
            this.LblCnte.BackColor = System.Drawing.Color.Transparent;
            this.LblCnte.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblCnte.ForeColor = System.Drawing.Color.White;
            this.LblCnte.Location = new System.Drawing.Point(1, 20);
            this.LblCnte.Name = "LblCnte";
            this.LblCnte.Size = new System.Drawing.Size(47, 20);
            this.LblCnte.TabIndex = 99;
            this.LblCnte.Text = "Cnte";
            // 
            // Clave
            // 
            this.Clave.Frozen = true;
            this.Clave.HeaderText = "Clave";
            this.Clave.Name = "Clave";
            this.Clave.ReadOnly = true;
            this.Clave.Width = 90;
            // 
            // LblPedido
            // 
            this.LblPedido.AutoSize = true;
            this.LblPedido.BackColor = System.Drawing.Color.Transparent;
            this.LblPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblPedido.ForeColor = System.Drawing.Color.White;
            this.LblPedido.Location = new System.Drawing.Point(1, 0);
            this.LblPedido.Name = "LblPedido";
            this.LblPedido.Size = new System.Drawing.Size(77, 20);
            this.LblPedido.TabIndex = 98;
            this.LblPedido.Text = "PEDIDO";
            // 
            // Importe
            // 
            this.Importe.HeaderText = "Importe";
            this.Importe.Name = "Importe";
            this.Importe.ReadOnly = true;
            this.Importe.Width = 70;
            // 
            // LblFecha
            // 
            this.LblFecha.AutoSize = true;
            this.LblFecha.BackColor = System.Drawing.Color.Transparent;
            this.LblFecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblFecha.ForeColor = System.Drawing.Color.White;
            this.LblFecha.Location = new System.Drawing.Point(658, -3);
            this.LblFecha.Name = "LblFecha";
            this.LblFecha.Size = new System.Drawing.Size(23, 13);
            this.LblFecha.TabIndex = 96;
            this.LblFecha.Text = "....";
            // 
            // DGDatos
            // 
            this.DGDatos.AllowUserToAddRows = false;
            this.DGDatos.AllowUserToDeleteRows = false;
            this.DGDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGDatos.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Clave,
            this.Nombre,
            this.Cajas,
            this.Importe});
            this.DGDatos.Location = new System.Drawing.Point(-4, 46);
            this.DGDatos.Name = "DGDatos";
            this.DGDatos.Size = new System.Drawing.Size(984, 214);
            this.DGDatos.TabIndex = 95;
            this.DGDatos.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGDatos_CellDoubleClick);
            this.DGDatos.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.DGDatos_DataError);
            // 
            // PBbtnsalir
            // 
            this.PBbtnsalir.Image = global::OrdenVentaExp.Properties.Resources.BtnSalir_Mini1;
            this.PBbtnsalir.Location = new System.Drawing.Point(943, 1);
            this.PBbtnsalir.Name = "PBbtnsalir";
            this.PBbtnsalir.Size = new System.Drawing.Size(30, 30);
            this.PBbtnsalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.PBbtnsalir.TabIndex = 103;
            this.PBbtnsalir.TabStop = false;
            this.PBbtnsalir.Visible = false;
            this.PBbtnsalir.Click += new System.EventHandler(this.PBbtnsalir_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BtnSave.Image = global::OrdenVentaExp.Properties.Resources.BtnGuardarMini;
            this.BtnSave.Location = new System.Drawing.Point(889, -3);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(35, 34);
            this.BtnSave.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.BtnSave.TabIndex = 97;
            this.BtnSave.TabStop = false;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // Frm_Faltante
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(976, 256);
            this.ControlBox = false;
            this.Controls.Add(this.PBbtnsalir);
            this.Controls.Add(this.LblAnio);
            this.Controls.Add(this.LblSemana);
            this.Controls.Add(this.LblNomCnte);
            this.Controls.Add(this.LblCnte);
            this.Controls.Add(this.LblPedido);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.LblFecha);
            this.Controls.Add(this.DGDatos);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Frm_Faltante";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "4.11   P R O D U C T O  F A L T A N T E";
            this.Load += new System.EventHandler(this.Frm_Faltante_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGDatos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BtnSave)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridViewTextBoxColumn Cajas;
        private System.Windows.Forms.PictureBox PBbtnsalir;
        private System.Windows.Forms.DataGridViewTextBoxColumn Nombre;
        private System.Windows.Forms.Label LblAnio;
        private System.Windows.Forms.Label LblSemana;
        private System.Windows.Forms.Label LblNomCnte;
        private System.Windows.Forms.Label LblCnte;
        private System.Windows.Forms.DataGridViewTextBoxColumn Clave;
        private System.Windows.Forms.Label LblPedido;
        private System.Windows.Forms.PictureBox BtnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn Importe;
        private System.Windows.Forms.Label LblFecha;
        private System.Windows.Forms.DataGridView DGDatos;
    }
}