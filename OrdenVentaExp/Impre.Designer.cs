namespace OrdenVentaExp
{
    partial class Impre
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Impre));
            this.PBbtnguardar = new System.Windows.Forms.PictureBox();
            this.PBbtnsalir = new System.Windows.Forms.PictureBox();
            this.CBselecimp = new System.Windows.Forms.ComboBox();
            this.lbselimp = new System.Windows.Forms.Label();
            this.lbimprdete = new System.Windows.Forms.Label();
            this.txtimpdeter = new System.Windows.Forms.TextBox();
            this.lbcab = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnguardar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).BeginInit();
            this.SuspendLayout();
            // 
            // PBbtnguardar
            // 
            this.PBbtnguardar.Image = global::OrdenVentaExp.Properties.Resources.save;
            this.PBbtnguardar.Location = new System.Drawing.Point(98, 217);
            this.PBbtnguardar.Name = "PBbtnguardar";
            this.PBbtnguardar.Size = new System.Drawing.Size(59, 55);
            this.PBbtnguardar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PBbtnguardar.TabIndex = 104;
            this.PBbtnguardar.TabStop = false;
            this.PBbtnguardar.Click += new System.EventHandler(this.PBbtnguardar_Click);
            // 
            // PBbtnsalir
            // 
            this.PBbtnsalir.Image = global::OrdenVentaExp.Properties.Resources._out;
            this.PBbtnsalir.Location = new System.Drawing.Point(312, 217);
            this.PBbtnsalir.Name = "PBbtnsalir";
            this.PBbtnsalir.Size = new System.Drawing.Size(59, 55);
            this.PBbtnsalir.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PBbtnsalir.TabIndex = 103;
            this.PBbtnsalir.TabStop = false;
            this.PBbtnsalir.Click += new System.EventHandler(this.PBbtnsalir_Click);
            // 
            // CBselecimp
            // 
            this.CBselecimp.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CBselecimp.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.CBselecimp.FormattingEnabled = true;
            this.CBselecimp.Location = new System.Drawing.Point(59, 158);
            this.CBselecimp.Name = "CBselecimp";
            this.CBselecimp.Size = new System.Drawing.Size(351, 24);
            this.CBselecimp.TabIndex = 102;
            this.CBselecimp.SelectionChangeCommitted += new System.EventHandler(this.CBselecimp_SelectionChangeCommitted);
            this.CBselecimp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CBselecimp_KeyDown);
            // 
            // lbselimp
            // 
            this.lbselimp.AutoSize = true;
            this.lbselimp.BackColor = System.Drawing.Color.Transparent;
            this.lbselimp.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbselimp.ForeColor = System.Drawing.SystemColors.Control;
            this.lbselimp.Location = new System.Drawing.Point(141, 117);
            this.lbselimp.Name = "lbselimp";
            this.lbselimp.Size = new System.Drawing.Size(186, 20);
            this.lbselimp.TabIndex = 101;
            this.lbselimp.Text = "Seleccionar impresora";
            // 
            // lbimprdete
            // 
            this.lbimprdete.AutoSize = true;
            this.lbimprdete.BackColor = System.Drawing.Color.Transparent;
            this.lbimprdete.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbimprdete.ForeColor = System.Drawing.SystemColors.Control;
            this.lbimprdete.Location = new System.Drawing.Point(123, 39);
            this.lbimprdete.Name = "lbimprdete";
            this.lbimprdete.Size = new System.Drawing.Size(223, 20);
            this.lbimprdete.TabIndex = 100;
            this.lbimprdete.Text = "Impresora predeterminada";
            // 
            // txtimpdeter
            // 
            this.txtimpdeter.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtimpdeter.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtimpdeter.Location = new System.Drawing.Point(108, 73);
            this.txtimpdeter.MaxLength = 6;
            this.txtimpdeter.Name = "txtimpdeter";
            this.txtimpdeter.Size = new System.Drawing.Size(253, 27);
            this.txtimpdeter.TabIndex = 99;
            // 
            // lbcab
            // 
            this.lbcab.AutoSize = true;
            this.lbcab.BackColor = System.Drawing.Color.Transparent;
            this.lbcab.Font = new System.Drawing.Font("Lucida Sans Unicode", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbcab.ForeColor = System.Drawing.Color.White;
            this.lbcab.Location = new System.Drawing.Point(28, 9);
            this.lbcab.Name = "lbcab";
            this.lbcab.Size = new System.Drawing.Size(413, 20);
            this.lbcab.TabIndex = 98;
            this.lbcab.Text = "CONFIGURAR IMPRESORA  COMO PREDETERMINADA";
            // 
            // Impre
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(469, 284);
            this.Controls.Add(this.PBbtnguardar);
            this.Controls.Add(this.PBbtnsalir);
            this.Controls.Add(this.CBselecimp);
            this.Controls.Add(this.lbselimp);
            this.Controls.Add(this.lbimprdete);
            this.Controls.Add(this.txtimpdeter);
            this.Controls.Add(this.lbcab);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Impre";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impre";
            this.Load += new System.EventHandler(this.Impre_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnguardar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PBbtnsalir)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox PBbtnguardar;
        private System.Windows.Forms.PictureBox PBbtnsalir;
        private System.Windows.Forms.ComboBox CBselecimp;
        private System.Windows.Forms.Label lbselimp;
        private System.Windows.Forms.Label lbimprdete;
        private System.Windows.Forms.TextBox txtimpdeter;
        private System.Windows.Forms.Label lbcab;
    }
}