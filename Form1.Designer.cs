namespace ExcelNPOI
{
    partial class Form1
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
            this.btnGerarCatalogo = new System.Windows.Forms.Button();
            this.saveFileDialogCatalogo = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // btnGerarCatalogo
            // 
            this.btnGerarCatalogo.Location = new System.Drawing.Point(121, 25);
            this.btnGerarCatalogo.Name = "btnGerarCatalogo";
            this.btnGerarCatalogo.Size = new System.Drawing.Size(378, 23);
            this.btnGerarCatalogo.TabIndex = 0;
            this.btnGerarCatalogo.Text = "Gerar excel";
            this.btnGerarCatalogo.UseVisualStyleBackColor = true;
            this.btnGerarCatalogo.Click += new System.EventHandler(this.btnGerarCatalogo_Click);
            // 
            // saveFileDialogCatalogo
            // 
            this.saveFileDialogCatalogo.Filter = ".xls Files (*.xls)|*.xls";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(605, 182);
            this.Controls.Add(this.btnGerarCatalogo);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnGerarCatalogo;
        private System.Windows.Forms.SaveFileDialog saveFileDialogCatalogo;
    }
}

