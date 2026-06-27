namespace OrdenVentaExp
{
    partial class FrmCompleExp
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
            this.components = new System.ComponentModel.Container();
            this.LblDire = new System.Windows.Forms.Label();
            this.LblDes = new System.Windows.Forms.Label();
            this.DgDatos = new System.Windows.Forms.DataGridView();
            this.lbtil = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.PbxExcel = new System.Windows.Forms.PictureBox();
            this.PbActFracc = new System.Windows.Forms.PictureBox();
            this.PbActu = new System.Windows.Forms.PictureBox();
            this.PBbtnsalir = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.DgDatos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbxExcel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbActFracc)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbActu)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).BeginInit();
            this.SuspendLayout();
            // 
            // LblDire
            // 
            this.LblDire.AutoSize = true;
            this.LblDire.BackColor = System.Drawing.Color.Transparent;
            this.LblDire.Font = new System.Drawing.Font("Lucida Sans Unicode", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblDire.ForeColor = System.Drawing.Color.White;
            this.LblDire.Location = new System.Drawing.Point(7, 73);
            this.LblDire.Name = "LblDire";
            this.LblDire.Size = new System.Drawing.Size(55, 15);
            this.LblDire.TabIndex = 120;
            this.LblDire.Text = "Direccón";
            // 
            // LblDes
            // 
            this.LblDes.AutoSize = true;
            this.LblDes.BackColor = System.Drawing.Color.Transparent;
            this.LblDes.Font = new System.Drawing.Font("Lucida Sans Unicode", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblDes.ForeColor = System.Drawing.Color.White;
            this.LblDes.Location = new System.Drawing.Point(7, 53);
            this.LblDes.Name = "LblDes";
            this.LblDes.Size = new System.Drawing.Size(72, 15);
            this.LblDes.TabIndex = 119;
            this.LblDes.Text = "Destinatario";
            // 
            // DgDatos
            // 
            this.DgDatos.AllowUserToAddRows = false;
            this.DgDatos.AllowUserToDeleteRows = false;
            this.DgDatos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgDatos.Location = new System.Drawing.Point(2, 96);
            this.DgDatos.Name = "DgDatos";
            this.DgDatos.Size = new System.Drawing.Size(1102, 391);
            this.DgDatos.TabIndex = 114;
            // 
            // lbtil
            // 
            this.lbtil.AutoSize = true;
            this.lbtil.BackColor = System.Drawing.Color.Transparent;
            this.lbtil.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbtil.ForeColor = System.Drawing.Color.White;
            this.lbtil.Location = new System.Drawing.Point(7, 21);
            this.lbtil.Name = "lbtil";
            this.lbtil.Size = new System.Drawing.Size(277, 20);
            this.lbtil.TabIndex = 113;
            this.lbtil.Text = "COMPLEMENTO DE  EXPORTACIÓN";
            // 
            // PbxExcel
            // 
            this.PbxExcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.PbxExcel.Image = global::OrdenVentaExp.Properties.Resources.Excel;
            this.PbxExcel.Location = new System.Drawing.Point(873, 12);
            this.PbxExcel.Name = "PbxExcel";
            this.PbxExcel.Size = new System.Drawing.Size(34, 34);
            this.PbxExcel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.PbxExcel.TabIndex = 118;
            this.PbxExcel.TabStop = false;
            this.toolTip1.SetToolTip(this.PbxExcel, "Información Carta Porte");
            this.PbxExcel.Click += new System.EventHandler(this.PbxExcel_Click);
            // 
            // PbActFracc
            // 
            this.PbActFracc.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.PbActFracc.Image = global::OrdenVentaExp.Properties.Resources.actualizar_mini;
            this.PbActFracc.Location = new System.Drawing.Point(499, 12);
            this.PbActFracc.Name = "PbActFracc";
            this.PbActFracc.Size = new System.Drawing.Size(24, 24);
            this.PbActFracc.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.PbActFracc.TabIndex = 117;
            this.PbActFracc.TabStop = false;
            this.toolTip1.SetToolTip(this.PbActFracc, "Actualiza La Fraccion Alancelaria del Producto");
            this.PbActFracc.Visible = false;
            // 
            // PbActu
            // 
            this.PbActu.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.PbActu.Image = global::OrdenVentaExp.Properties.Resources.actualizar_mini;
            this.PbActu.Location = new System.Drawing.Point(682, 12);
            this.PbActu.Name = "PbActu";
            this.PbActu.Size = new System.Drawing.Size(24, 24);
            this.PbActu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.PbActu.TabIndex = 116;
            this.PbActu.TabStop = false;
            this.toolTip1.SetToolTip(this.PbActu, "Actualiza El Peso del Producto");
            this.PbActu.Visible = false;
            // 
            // PBbtnsalir
            // 
            this.PBbtnsalir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.PBbtnsalir.Image = global::OrdenVentaExp.Properties.Resources.BtnSalir_Mini;
            this.PBbtnsalir.Location = new System.Drawing.Point(1069, 12);
            this.PBbtnsalir.Name = "PBbtnsalir";
            this.PBbtnsalir.Size = new System.Drawing.Size(30, 30);
            this.PBbtnsalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.PBbtnsalir.TabIndex = 115;
            this.PBbtnsalir.TabStop = false;
            this.toolTip1.SetToolTip(this.PBbtnsalir, "Salir de la Pantalla");
            this.PBbtnsalir.Click += new System.EventHandler(this.PBbtnsalir_Click);
            // 
            // FrmCompleExp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkGray;
            this.ClientSize = new System.Drawing.Size(1113, 497);
            this.ControlBox = false;
            this.Controls.Add(this.LblDire);
            this.Controls.Add(this.LblDes);
            this.Controls.Add(this.PbxExcel);
            this.Controls.Add(this.PbActFracc);
            this.Controls.Add(this.PbActu);
            this.Controls.Add(this.PBbtnsalir);
            this.Controls.Add(this.DgDatos);
            this.Controls.Add(this.lbtil);
            this.Name = "FrmCompleExp";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CARTA PORTE";
            this.Load += new System.EventHandler(this.FrmCompleExp_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DgDatos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbxExcel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbActFracc)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PbActu)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LblDire;
        private System.Windows.Forms.Label LblDes;
        private System.Windows.Forms.PictureBox PbxExcel;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.PictureBox PbActFracc;
        private System.Windows.Forms.PictureBox PbActu;
        private System.Windows.Forms.PictureBox PBbtnsalir;
        private System.Windows.Forms.DataGridView DgDatos;
        private System.Windows.Forms.Label lbtil;
    }
}