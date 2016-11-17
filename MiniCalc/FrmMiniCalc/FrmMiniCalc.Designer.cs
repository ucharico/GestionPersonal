namespace FrmMiniCalc
{
    partial class FrmMiniCalc
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMiniCalc));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.NumberA = new System.Windows.Forms.NumericUpDown();
            this.NumberB = new System.Windows.Forms.NumericUpDown();
            this.AddRButton = new System.Windows.Forms.RadioButton();
            this.SubtractRButton = new System.Windows.Forms.RadioButton();
            this.EqualsButton = new System.Windows.Forms.Button();
            this.Result = new System.Windows.Forms.Label();
            this.CmdSalir = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.NumberA)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.NumberB)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Ingrese el Primer Número";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(263, 97);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(139, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Ingrese el Segundo Número";
            // 
            // NumberA
            // 
            this.NumberA.Location = new System.Drawing.Point(30, 114);
            this.NumberA.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.NumberA.Name = "NumberA";
            this.NumberA.Size = new System.Drawing.Size(120, 20);
            this.NumberA.TabIndex = 4;
            // 
            // NumberB
            // 
            this.NumberB.Location = new System.Drawing.Point(266, 114);
            this.NumberB.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.NumberB.Name = "NumberB";
            this.NumberB.Size = new System.Drawing.Size(120, 20);
            this.NumberB.TabIndex = 5;
            // 
            // AddRButton
            // 
            this.AddRButton.AutoSize = true;
            this.AddRButton.Checked = true;
            this.AddRButton.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AddRButton.Location = new System.Drawing.Point(176, 89);
            this.AddRButton.Name = "AddRButton";
            this.AddRButton.Size = new System.Drawing.Size(37, 27);
            this.AddRButton.TabIndex = 6;
            this.AddRButton.TabStop = true;
            this.AddRButton.Text = "+";
            this.AddRButton.UseVisualStyleBackColor = true;
            this.AddRButton.CheckedChanged += new System.EventHandler(this.AddRButton_CheckedChanged);
            // 
            // SubtractRButton
            // 
            this.SubtractRButton.AutoSize = true;
            this.SubtractRButton.Font = new System.Drawing.Font("Calibri", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SubtractRButton.Location = new System.Drawing.Point(176, 112);
            this.SubtractRButton.Name = "SubtractRButton";
            this.SubtractRButton.Size = new System.Drawing.Size(41, 37);
            this.SubtractRButton.TabIndex = 7;
            this.SubtractRButton.Text = "-";
            this.SubtractRButton.UseVisualStyleBackColor = true;
            this.SubtractRButton.CheckedChanged += new System.EventHandler(this.SubtractRButton_CheckedChanged);
            // 
            // EqualsButton
            // 
            this.EqualsButton.Location = new System.Drawing.Point(425, 108);
            this.EqualsButton.Name = "EqualsButton";
            this.EqualsButton.Size = new System.Drawing.Size(73, 26);
            this.EqualsButton.TabIndex = 8;
            this.EqualsButton.Text = "=";
            this.EqualsButton.UseVisualStyleBackColor = true;
            this.EqualsButton.Click += new System.EventHandler(this.EqualsButton_Click);
            // 
            // Result
            // 
            this.Result.AutoSize = true;
            this.Result.Font = new System.Drawing.Font("Calibri", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Result.Location = new System.Drawing.Point(516, 109);
            this.Result.Name = "Result";
            this.Result.Size = new System.Drawing.Size(61, 23);
            this.Result.TabIndex = 9;
            this.Result.Text = "Result";
            // 
            // CmdSalir
            // 
            this.CmdSalir.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CmdSalir.Image = ((System.Drawing.Image)(resources.GetObject("CmdSalir.Image")));
            this.CmdSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.CmdSalir.Location = new System.Drawing.Point(553, 21);
            this.CmdSalir.Name = "CmdSalir";
            this.CmdSalir.Size = new System.Drawing.Size(85, 47);
            this.CmdSalir.TabIndex = 10;
            this.CmdSalir.Text = "&Cerrar";
            this.CmdSalir.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CmdSalir.UseVisualStyleBackColor = true;
            this.CmdSalir.Click += new System.EventHandler(this.CmdSalir_Click);
            // 
            // FrmMiniCalc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(665, 192);
            this.Controls.Add(this.CmdSalir);
            this.Controls.Add(this.Result);
            this.Controls.Add(this.EqualsButton);
            this.Controls.Add(this.SubtractRButton);
            this.Controls.Add(this.AddRButton);
            this.Controls.Add(this.NumberB);
            this.Controls.Add(this.NumberA);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "FrmMiniCalc";
            this.Text = "MiniCalc";
            this.Load += new System.EventHandler(this.FrmMiniCalc_Load);
            ((System.ComponentModel.ISupportInitialize)(this.NumberA)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.NumberB)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown NumberA;
        private System.Windows.Forms.NumericUpDown NumberB;
        private System.Windows.Forms.RadioButton AddRButton;
        private System.Windows.Forms.RadioButton SubtractRButton;
        private System.Windows.Forms.Button EqualsButton;
        private System.Windows.Forms.Label Result;
        private System.Windows.Forms.Button CmdSalir;
    }
}

