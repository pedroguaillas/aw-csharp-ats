namespace WinAppATS
{
    partial class FormRegistroAnulados
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgvAnulados = new System.Windows.Forms.DataGridView();
            this.tipoComprobante = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.establecimiento = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.puntoEmision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.secuencialInicio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.secuencialFin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.autorizacion = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAnulados)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvAnulados
            // 
            this.dgvAnulados.AllowUserToAddRows = false;
            this.dgvAnulados.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dgvAnulados.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.DodgerBlue;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Franklin Gothic Book", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvAnulados.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvAnulados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAnulados.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.tipoComprobante,
            this.establecimiento,
            this.puntoEmision,
            this.secuencialInicio,
            this.secuencialFin,
            this.autorizacion});
            this.dgvAnulados.EnableHeadersVisualStyles = false;
            this.dgvAnulados.Location = new System.Drawing.Point(230, 170);
            this.dgvAnulados.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dgvAnulados.Name = "dgvAnulados";
            this.dgvAnulados.Size = new System.Drawing.Size(990, 600);
            this.dgvAnulados.TabIndex = 2;
            this.dgvAnulados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DgvAnulados_KeyDown);
            // 
            // tipoComprobante
            // 
            this.tipoComprobante.HeaderText = "Código tipo comprobante anulado";
            this.tipoComprobante.MaxInputLength = 2;
            this.tipoComprobante.Name = "tipoComprobante";
            // 
            // establecimiento
            // 
            this.establecimiento.HeaderText = "Establecimiento";
            this.establecimiento.MaxInputLength = 3;
            this.establecimiento.Name = "establecimiento";
            // 
            // puntoEmision
            // 
            this.puntoEmision.HeaderText = "Punto Emisión";
            this.puntoEmision.MaxInputLength = 3;
            this.puntoEmision.Name = "puntoEmision";
            // 
            // secuencialInicio
            // 
            this.secuencialInicio.HeaderText = "Secuencial inicio";
            this.secuencialInicio.MaxInputLength = 9;
            this.secuencialInicio.Name = "secuencialInicio";
            // 
            // secuencialFin
            // 
            this.secuencialFin.HeaderText = "Secuencial Fin";
            this.secuencialFin.MaxInputLength = 9;
            this.secuencialFin.Name = "secuencialFin";
            // 
            // autorizacion
            // 
            this.autorizacion.HeaderText = "Autorización";
            this.autorizacion.MaxInputLength = 49;
            this.autorizacion.Name = "autorizacion";
            this.autorizacion.Width = 200;
            // 
            // btnGuardar
            // 
            this.btnGuardar.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnGuardar.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnGuardar.FlatAppearance.BorderSize = 0;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.Font = new System.Drawing.Font("Franklin Gothic Book", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.Location = new System.Drawing.Point(230, 116);
            this.btnGuardar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(100, 28);
            this.btnGuardar.TabIndex = 3;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.BtnGuardar_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnAdd.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnAdd.FlatAppearance.BorderSize = 0;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAdd.Font = new System.Drawing.Font("Franklin Gothic Book", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAdd.ForeColor = System.Drawing.Color.White;
            this.btnAdd.Location = new System.Drawing.Point(355, 116);
            this.btnAdd.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(100, 28);
            this.btnAdd.TabIndex = 4;
            this.btnAdd.Text = "Nuevo";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // FormRegistroAnulados
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(32)))), ((int)(((byte)(40)))));
            this.ClientSize = new System.Drawing.Size(1440, 921);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.dgvAnulados);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "FormRegistroAnulados";
            this.Text = "FormRegistroAnulados";
            this.Load += new System.EventHandler(this.FormRegistroAnulados_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvAnulados)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvAnulados;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.DataGridViewTextBoxColumn tipoComprobante;
        private System.Windows.Forms.DataGridViewTextBoxColumn establecimiento;
        private System.Windows.Forms.DataGridViewTextBoxColumn puntoEmision;
        private System.Windows.Forms.DataGridViewTextBoxColumn secuencialInicio;
        private System.Windows.Forms.DataGridViewTextBoxColumn secuencialFin;
        private System.Windows.Forms.DataGridViewTextBoxColumn autorizacion;
    }
}