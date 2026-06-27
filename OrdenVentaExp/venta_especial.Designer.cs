namespace OrdenVentaExp
{
    partial class venta_especial
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(venta_especial));
            this.txtclaveprod = new System.Windows.Forms.TextBox();
            this.lbprod = new System.Windows.Forms.Label();
            this.txtdescripcion = new System.Windows.Forms.TextBox();
            this.bntsave = new System.Windows.Forms.Button();
            this.btnexit = new System.Windows.Forms.Button();
            this.lbnumchar = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtclaveprod
            // 
            this.txtclaveprod.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtclaveprod.Enabled = false;
            this.txtclaveprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtclaveprod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(43)))), ((int)(((byte)(136)))));
            this.txtclaveprod.Location = new System.Drawing.Point(82, 12);
            this.txtclaveprod.MaxLength = 10;
            this.txtclaveprod.Name = "txtclaveprod";
            this.txtclaveprod.Size = new System.Drawing.Size(437, 27);
            this.txtclaveprod.TabIndex = 149;
            // 
            // lbprod
            // 
            this.lbprod.AutoSize = true;
            this.lbprod.BackColor = System.Drawing.Color.Transparent;
            this.lbprod.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbprod.ForeColor = System.Drawing.SystemColors.Control;
            this.lbprod.Location = new System.Drawing.Point(12, 15);
            this.lbprod.Name = "lbprod";
            this.lbprod.Size = new System.Drawing.Size(64, 16);
            this.lbprod.TabIndex = 148;
            this.lbprod.Text = "Producto";
            // 
            // txtdescripcion
            // 
            this.txtdescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtdescripcion.Location = new System.Drawing.Point(15, 45);
            this.txtdescripcion.MaxLength = 515;
            this.txtdescripcion.Multiline = true;
            this.txtdescripcion.Name = "txtdescripcion";
            this.txtdescripcion.Size = new System.Drawing.Size(504, 140);
            this.txtdescripcion.TabIndex = 150;
            this.txtdescripcion.TextChanged += new System.EventHandler(this.txtdescripcion_TextChanged);
            this.txtdescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtdescripcion_KeyPress);
            // 
            // bntsave
            // 
            this.bntsave.BackgroundImage = global::OrdenVentaExp.Properties.Resources.save;
            this.bntsave.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bntsave.Location = new System.Drawing.Point(15, 191);
            this.bntsave.Name = "bntsave";
            this.bntsave.Size = new System.Drawing.Size(41, 38);
            this.bntsave.TabIndex = 151;
            this.bntsave.UseVisualStyleBackColor = true;
            this.bntsave.Click += new System.EventHandler(this.bntsave_Click);
            // 
            // btnexit
            // 
            this.btnexit.BackgroundImage = global::OrdenVentaExp.Properties.Resources.salir;
            this.btnexit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnexit.Location = new System.Drawing.Point(478, 191);
            this.btnexit.Name = "btnexit";
            this.btnexit.Size = new System.Drawing.Size(41, 38);
            this.btnexit.TabIndex = 152;
            this.btnexit.UseVisualStyleBackColor = true;
            this.btnexit.Click += new System.EventHandler(this.btnexit_Click);
            // 
            // lbnumchar
            // 
            this.lbnumchar.AutoSize = true;
            this.lbnumchar.BackColor = System.Drawing.Color.Transparent;
            this.lbnumchar.Font = new System.Drawing.Font("Lucida Sans Unicode", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbnumchar.ForeColor = System.Drawing.Color.Black;
            this.lbnumchar.Location = new System.Drawing.Point(151, 210);
            this.lbnumchar.Name = "lbnumchar";
            this.lbnumchar.Size = new System.Drawing.Size(228, 16);
            this.lbnumchar.TabIndex = 153;
            this.lbnumchar.Text = "Quedan 483 caracteres por escribir";
            // 
            // venta_especial
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(531, 235);
            this.ControlBox = false;
            this.Controls.Add(this.lbnumchar);
            this.Controls.Add(this.btnexit);
            this.Controls.Add(this.bntsave);
            this.Controls.Add(this.txtdescripcion);
            this.Controls.Add(this.txtclaveprod);
            this.Controls.Add(this.lbprod);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "venta_especial";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Observación Especial";
            this.Load += new System.EventHandler(this.venta_especial_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtclaveprod;
        private System.Windows.Forms.Label lbprod;
        private System.Windows.Forms.TextBox txtdescripcion;
        private System.Windows.Forms.Button bntsave;
        private System.Windows.Forms.Button btnexit;
        private System.Windows.Forms.Label lbnumchar;
    }
}