namespace OrdenVentaExp
{
    partial class CajaTrailer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CajaTrailer));
            this.PBbtnguardar = new System.Windows.Forms.PictureBox();
            this.PBbtnsalir = new System.Windows.Forms.PictureBox();
            this.txtct = new System.Windows.Forms.TextBox();
            this.lbcaja = new System.Windows.Forms.Label();
            this.txtfolio = new System.Windows.Forms.TextBox();
            this.lbfolio = new System.Windows.Forms.Label();
            this.lbcab = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnguardar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).BeginInit();
            this.SuspendLayout();
            // 
            // PBbtnguardar
            // 
            this.PBbtnguardar.Image = global::OrdenVentaExp.Properties.Resources.save;
            this.PBbtnguardar.Location = new System.Drawing.Point(61, 185);
            this.PBbtnguardar.Name = "PBbtnguardar";
            this.PBbtnguardar.Size = new System.Drawing.Size(59, 55);
            this.PBbtnguardar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PBbtnguardar.TabIndex = 102;
            this.PBbtnguardar.TabStop = false;
            this.PBbtnguardar.Click += new System.EventHandler(this.PBbtnguardar_Click);
            // 
            // PBbtnsalir
            // 
            this.PBbtnsalir.Image = global::OrdenVentaExp.Properties.Resources._out;
            this.PBbtnsalir.Location = new System.Drawing.Point(275, 185);
            this.PBbtnsalir.Name = "PBbtnsalir";
            this.PBbtnsalir.Size = new System.Drawing.Size(59, 55);
            this.PBbtnsalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PBbtnsalir.TabIndex = 101;
            this.PBbtnsalir.TabStop = false;
            this.PBbtnsalir.Click += new System.EventHandler(this.PBbtnsalir_Click);
            // 
            // txtct
            // 
            this.txtct.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtct.Location = new System.Drawing.Point(205, 122);
            this.txtct.MaxLength = 8;
            this.txtct.Name = "txtct";
            this.txtct.Size = new System.Drawing.Size(102, 27);
            this.txtct.TabIndex = 100;
            // 
            // lbcaja
            // 
            this.lbcaja.AutoSize = true;
            this.lbcaja.BackColor = System.Drawing.Color.Transparent;
            this.lbcaja.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcaja.ForeColor = System.Drawing.SystemColors.Control;
            this.lbcaja.Location = new System.Drawing.Point(97, 121);
            this.lbcaja.Name = "lbcaja";
            this.lbcaja.Size = new System.Drawing.Size(102, 20);
            this.lbcaja.TabIndex = 99;
            this.lbcaja.Text = "Caja trailer:";
            // 
            // txtfolio
            // 
            this.txtfolio.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtfolio.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtfolio.Location = new System.Drawing.Point(205, 68);
            this.txtfolio.MaxLength = 6;
            this.txtfolio.Name = "txtfolio";
            this.txtfolio.Size = new System.Drawing.Size(102, 27);
            this.txtfolio.TabIndex = 98;
            this.txtfolio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtfolio_KeyPress);
            // 
            // lbfolio
            // 
            this.lbfolio.AutoSize = true;
            this.lbfolio.BackColor = System.Drawing.Color.Transparent;
            this.lbfolio.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbfolio.ForeColor = System.Drawing.SystemColors.Control;
            this.lbfolio.Location = new System.Drawing.Point(146, 67);
            this.lbfolio.Name = "lbfolio";
            this.lbfolio.Size = new System.Drawing.Size(53, 20);
            this.lbfolio.TabIndex = 97;
            this.lbfolio.Text = "Folio:";
            // 
            // lbcab
            // 
            this.lbcab.AutoSize = true;
            this.lbcab.BackColor = System.Drawing.Color.Transparent;
            this.lbcab.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcab.ForeColor = System.Drawing.Color.White;
            this.lbcab.Location = new System.Drawing.Point(21, 14);
            this.lbcab.Name = "lbcab";
            this.lbcab.Size = new System.Drawing.Size(362, 20);
            this.lbcab.TabIndex = 96;
            this.lbcab.Text = "ASIGNAR O MODIFICAR LA CAJA DEL TRAILER";
            // 
            // CajaTrailer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(405, 254);
            this.Controls.Add(this.PBbtnguardar);
            this.Controls.Add(this.PBbtnsalir);
            this.Controls.Add(this.txtct);
            this.Controls.Add(this.lbcaja);
            this.Controls.Add(this.txtfolio);
            this.Controls.Add(this.lbfolio);
            this.Controls.Add(this.lbcab);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CajaTrailer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CajaTrailer";
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnguardar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox PBbtnguardar;
        private System.Windows.Forms.PictureBox PBbtnsalir;
        private System.Windows.Forms.TextBox txtct;
        private System.Windows.Forms.Label lbcaja;
        private System.Windows.Forms.TextBox txtfolio;
        private System.Windows.Forms.Label lbfolio;
        private System.Windows.Forms.Label lbcab;
    }
}