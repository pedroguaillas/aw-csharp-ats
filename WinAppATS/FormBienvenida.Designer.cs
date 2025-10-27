namespace WinAppATS
{
    partial class FormBienvenida
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormBienvenida));
            this.btnExportar = new System.Windows.Forms.Button();
            this.Cliente = new System.Windows.Forms.GroupBox();
            this.cbDeclaracionSemestral = new System.Windows.Forms.CheckBox();
            this.tbContabilidad = new System.Windows.Forms.TextBox();
            this.tbRepresentanteLegal = new System.Windows.Forms.TextBox();
            this.tbDiaDeclaracion = new System.Windows.Forms.TextBox();
            this.tbEstablecimiento = new System.Windows.Forms.TextBox();
            this.tbRuc = new System.Windows.Forms.TextBox();
            this.tbMes = new System.Windows.Forms.TextBox();
            this.tbAnio = new System.Windows.Forms.TextBox();
            this.tbRazonSocial = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbSuper = new System.Windows.Forms.TextBox();
            this.tbMrl = new System.Windows.Forms.TextBox();
            this.tbMt = new System.Windows.Forms.TextBox();
            this.tbIess2 = new System.Windows.Forms.TextBox();
            this.tbIess1 = new System.Windows.Forms.TextBox();
            this.tbSri = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnExportDeclaracionJson = new System.Windows.Forms.Button();
            this.btnDeclaracionRets = new System.Windows.Forms.Button();
            this.cbVentasElectronicas = new System.Windows.Forms.CheckBox();
            this.Cliente.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExportar
            // 
            this.btnExportar.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.btnExportar.FlatAppearance.BorderSize = 0;
            this.btnExportar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExportar.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportar.ForeColor = System.Drawing.Color.White;
            this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
            this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportar.Location = new System.Drawing.Point(395, 725);
            this.btnExportar.Margin = new System.Windows.Forms.Padding(4);
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.Size = new System.Drawing.Size(165, 43);
            this.btnExportar.TabIndex = 1;
            this.btnExportar.Text = "     Exportar ATS";
            this.btnExportar.UseVisualStyleBackColor = false;
            this.btnExportar.Click += new System.EventHandler(this.BtnExportar_Click);
            // 
            // Cliente
            // 
            this.Cliente.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Cliente.Controls.Add(this.cbVentasElectronicas);
            this.Cliente.Controls.Add(this.cbDeclaracionSemestral);
            this.Cliente.Controls.Add(this.tbContabilidad);
            this.Cliente.Controls.Add(this.tbRepresentanteLegal);
            this.Cliente.Controls.Add(this.tbDiaDeclaracion);
            this.Cliente.Controls.Add(this.tbEstablecimiento);
            this.Cliente.Controls.Add(this.tbRuc);
            this.Cliente.Controls.Add(this.tbMes);
            this.Cliente.Controls.Add(this.tbAnio);
            this.Cliente.Controls.Add(this.tbRazonSocial);
            this.Cliente.Controls.Add(this.label17);
            this.Cliente.Controls.Add(this.label12);
            this.Cliente.Controls.Add(this.label10);
            this.Cliente.Controls.Add(this.label6);
            this.Cliente.Controls.Add(this.label5);
            this.Cliente.Controls.Add(this.label4);
            this.Cliente.Controls.Add(this.label3);
            this.Cliente.Controls.Add(this.label2);
            this.Cliente.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Cliente.ForeColor = System.Drawing.Color.White;
            this.Cliente.Location = new System.Drawing.Point(87, 182);
            this.Cliente.Margin = new System.Windows.Forms.Padding(4);
            this.Cliente.Name = "Cliente";
            this.Cliente.Padding = new System.Windows.Forms.Padding(4);
            this.Cliente.Size = new System.Drawing.Size(1229, 319);
            this.Cliente.TabIndex = 2;
            this.Cliente.TabStop = false;
            this.Cliente.Text = "Informante";
            // 
            // cbDeclaracionSemestral
            // 
            this.cbDeclaracionSemestral.AutoSize = true;
            this.cbDeclaracionSemestral.Location = new System.Drawing.Point(626, 255);
            this.cbDeclaracionSemestral.Name = "cbDeclaracionSemestral";
            this.cbDeclaracionSemestral.Size = new System.Drawing.Size(223, 29);
            this.cbDeclaracionSemestral.TabIndex = 9;
            this.cbDeclaracionSemestral.Text = "Declaración semestral";
            this.cbDeclaracionSemestral.UseVisualStyleBackColor = true;
            // 
            // tbContabilidad
            // 
            this.tbContabilidad.Location = new System.Drawing.Point(857, 151);
            this.tbContabilidad.Margin = new System.Windows.Forms.Padding(4);
            this.tbContabilidad.MaxLength = 2;
            this.tbContabilidad.Name = "tbContabilidad";
            this.tbContabilidad.Size = new System.Drawing.Size(332, 30);
            this.tbContabilidad.TabIndex = 7;
            this.tbContabilidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TbContabilidad_KeyPress);
            // 
            // tbRepresentanteLegal
            // 
            this.tbRepresentanteLegal.Location = new System.Drawing.Point(857, 101);
            this.tbRepresentanteLegal.Margin = new System.Windows.Forms.Padding(4);
            this.tbRepresentanteLegal.MaxLength = 300;
            this.tbRepresentanteLegal.Name = "tbRepresentanteLegal";
            this.tbRepresentanteLegal.Size = new System.Drawing.Size(332, 30);
            this.tbRepresentanteLegal.TabIndex = 7;
            this.tbRepresentanteLegal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TbRepresentanteLegal_KeyPress);
            // 
            // tbDiaDeclaracion
            // 
            this.tbDiaDeclaracion.Location = new System.Drawing.Point(857, 50);
            this.tbDiaDeclaracion.Margin = new System.Windows.Forms.Padding(4);
            this.tbDiaDeclaracion.MaxLength = 2;
            this.tbDiaDeclaracion.Name = "tbDiaDeclaracion";
            this.tbDiaDeclaracion.Size = new System.Drawing.Size(332, 30);
            this.tbDiaDeclaracion.TabIndex = 7;
            this.tbDiaDeclaracion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TbDiaDeclaracion_KeyPress);
            // 
            // tbEstablecimiento
            // 
            this.tbEstablecimiento.Location = new System.Drawing.Point(207, 254);
            this.tbEstablecimiento.Margin = new System.Windows.Forms.Padding(4);
            this.tbEstablecimiento.MaxLength = 3;
            this.tbEstablecimiento.Name = "tbEstablecimiento";
            this.tbEstablecimiento.Size = new System.Drawing.Size(332, 30);
            this.tbEstablecimiento.TabIndex = 7;
            this.tbEstablecimiento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TbEstablecimiento_KeyPress);
            // 
            // tbRuc
            // 
            this.tbRuc.Location = new System.Drawing.Point(207, 50);
            this.tbRuc.Margin = new System.Windows.Forms.Padding(4);
            this.tbRuc.MaxLength = 300;
            this.tbRuc.Name = "tbRuc";
            this.tbRuc.ReadOnly = true;
            this.tbRuc.Size = new System.Drawing.Size(332, 30);
            this.tbRuc.TabIndex = 7;
            // 
            // tbMes
            // 
            this.tbMes.Location = new System.Drawing.Point(207, 203);
            this.tbMes.Margin = new System.Windows.Forms.Padding(4);
            this.tbMes.MaxLength = 300;
            this.tbMes.Name = "tbMes";
            this.tbMes.ReadOnly = true;
            this.tbMes.Size = new System.Drawing.Size(332, 30);
            this.tbMes.TabIndex = 7;
            // 
            // tbAnio
            // 
            this.tbAnio.Location = new System.Drawing.Point(207, 152);
            this.tbAnio.Margin = new System.Windows.Forms.Padding(4);
            this.tbAnio.MaxLength = 300;
            this.tbAnio.Name = "tbAnio";
            this.tbAnio.ReadOnly = true;
            this.tbAnio.Size = new System.Drawing.Size(332, 30);
            this.tbAnio.TabIndex = 7;
            // 
            // tbRazonSocial
            // 
            this.tbRazonSocial.Location = new System.Drawing.Point(207, 101);
            this.tbRazonSocial.Margin = new System.Windows.Forms.Padding(4);
            this.tbRazonSocial.MaxLength = 300;
            this.tbRazonSocial.Name = "tbRazonSocial";
            this.tbRazonSocial.ReadOnly = true;
            this.tbRazonSocial.Size = new System.Drawing.Size(332, 30);
            this.tbRazonSocial.TabIndex = 7;
            this.tbRazonSocial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbRazonSocial_KeyDown);
            this.tbRazonSocial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TbRazonSocial_KeyPress);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(619, 154);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(199, 25);
            this.label17.TabIndex = 5;
            this.label17.Text = "Obligado contabilidad";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(619, 104);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(182, 25);
            this.label12.TabIndex = 5;
            this.label12.Text = "Representante legal";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(619, 53);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(145, 25);
            this.label10.TabIndex = 3;
            this.label10.Text = "Dia declaración";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(25, 257);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(147, 25);
            this.label6.TabIndex = 0;
            this.label6.Text = "Establecimiento";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(25, 206);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 25);
            this.label5.TabIndex = 0;
            this.label5.Text = "Mes";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 155);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 25);
            this.label4.TabIndex = 0;
            this.label4.Text = "Año";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 101);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(118, 25);
            this.label3.TabIndex = 0;
            this.label3.Text = "Razón social";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 50);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 25);
            this.label2.TabIndex = 0;
            this.label2.Text = "RUC";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.groupBox1.Controls.Add(this.tbSuper);
            this.groupBox1.Controls.Add(this.tbMrl);
            this.groupBox1.Controls.Add(this.tbMt);
            this.groupBox1.Controls.Add(this.tbIess2);
            this.groupBox1.Controls.Add(this.tbIess1);
            this.groupBox1.Controls.Add(this.tbSri);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(87, 539);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(1229, 131);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Claves";
            // 
            // tbSuper
            // 
            this.tbSuper.Location = new System.Drawing.Point(1056, 65);
            this.tbSuper.Margin = new System.Windows.Forms.Padding(4);
            this.tbSuper.MaxLength = 50;
            this.tbSuper.Name = "tbSuper";
            this.tbSuper.Size = new System.Drawing.Size(132, 30);
            this.tbSuper.TabIndex = 9;
            // 
            // tbMrl
            // 
            this.tbMrl.Location = new System.Drawing.Point(851, 65);
            this.tbMrl.Margin = new System.Windows.Forms.Padding(4);
            this.tbMrl.MaxLength = 50;
            this.tbMrl.Name = "tbMrl";
            this.tbMrl.Size = new System.Drawing.Size(132, 30);
            this.tbMrl.TabIndex = 9;
            this.tbMrl.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TbMrl_KeyUp);
            // 
            // tbMt
            // 
            this.tbMt.Location = new System.Drawing.Point(646, 68);
            this.tbMt.Margin = new System.Windows.Forms.Padding(4);
            this.tbMt.MaxLength = 50;
            this.tbMt.Name = "tbMt";
            this.tbMt.Size = new System.Drawing.Size(132, 30);
            this.tbMt.TabIndex = 9;
            this.tbMt.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TbMt_KeyUp);
            // 
            // tbIess2
            // 
            this.tbIess2.Location = new System.Drawing.Point(441, 68);
            this.tbIess2.Margin = new System.Windows.Forms.Padding(4);
            this.tbIess2.MaxLength = 50;
            this.tbIess2.Name = "tbIess2";
            this.tbIess2.Size = new System.Drawing.Size(132, 30);
            this.tbIess2.TabIndex = 9;
            this.tbIess2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TbIess2_KeyUp);
            // 
            // tbIess1
            // 
            this.tbIess1.Location = new System.Drawing.Point(236, 68);
            this.tbIess1.Margin = new System.Windows.Forms.Padding(4);
            this.tbIess1.MaxLength = 50;
            this.tbIess1.Name = "tbIess1";
            this.tbIess1.Size = new System.Drawing.Size(132, 30);
            this.tbIess1.TabIndex = 9;
            this.tbIess1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TbIess1_KeyUp);
            // 
            // tbSri
            // 
            this.tbSri.Location = new System.Drawing.Point(31, 68);
            this.tbSri.Margin = new System.Windows.Forms.Padding(4);
            this.tbSri.MaxLength = 50;
            this.tbSri.Name = "tbSri";
            this.tbSri.Size = new System.Drawing.Size(132, 30);
            this.tbSri.TabIndex = 9;
            this.tbSri.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TbSri_KeyUp);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(1051, 36);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(69, 25);
            this.label16.TabIndex = 8;
            this.label16.Text = "SUPER";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(846, 36);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(50, 25);
            this.label15.TabIndex = 8;
            this.label15.Text = "MRL";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(436, 36);
            this.label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(62, 25);
            this.label18.TabIndex = 6;
            this.label18.Text = "IESS2";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(641, 36);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(38, 25);
            this.label14.TabIndex = 7;
            this.label14.Text = "MT";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(231, 36);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(62, 25);
            this.label13.TabIndex = 6;
            this.label13.Text = "IESS1";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(32, 36);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(40, 25);
            this.label11.TabIndex = 4;
            this.label11.Text = "SRI";
            // 
            // btnGuardar
            // 
            this.btnGuardar.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.btnGuardar.FlatAppearance.BorderSize = 0;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
            this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGuardar.Location = new System.Drawing.Point(941, 725);
            this.btnGuardar.Margin = new System.Windows.Forms.Padding(4);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(165, 43);
            this.btnGuardar.TabIndex = 4;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.BtnGuardar_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(600, 26);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(152, 139);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // btnExportDeclaracionJson
            // 
            this.btnExportDeclaracionJson.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.btnExportDeclaracionJson.FlatAppearance.BorderSize = 0;
            this.btnExportDeclaracionJson.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExportDeclaracionJson.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportDeclaracionJson.ForeColor = System.Drawing.Color.White;
            this.btnExportDeclaracionJson.Image = ((System.Drawing.Image)(resources.GetObject("btnExportDeclaracionJson.Image")));
            this.btnExportDeclaracionJson.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnExportDeclaracionJson.Location = new System.Drawing.Point(587, 725);
            this.btnExportDeclaracionJson.Margin = new System.Windows.Forms.Padding(4);
            this.btnExportDeclaracionJson.Name = "btnExportDeclaracionJson";
            this.btnExportDeclaracionJson.Size = new System.Drawing.Size(150, 43);
            this.btnExportDeclaracionJson.TabIndex = 6;
            this.btnExportDeclaracionJson.Text = "  IVA JSON";
            this.btnExportDeclaracionJson.UseVisualStyleBackColor = false;
            this.btnExportDeclaracionJson.Click += new System.EventHandler(this.btnExportDeclaracionJson_Click);
            // 
            // btnDeclaracionRets
            // 
            this.btnDeclaracionRets.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.btnDeclaracionRets.FlatAppearance.BorderSize = 0;
            this.btnDeclaracionRets.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDeclaracionRets.Font = new System.Drawing.Font("Franklin Gothic Book", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeclaracionRets.ForeColor = System.Drawing.Color.White;
            this.btnDeclaracionRets.Image = ((System.Drawing.Image)(resources.GetObject("btnDeclaracionRets.Image")));
            this.btnDeclaracionRets.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDeclaracionRets.Location = new System.Drawing.Point(764, 725);
            this.btnDeclaracionRets.Margin = new System.Windows.Forms.Padding(4);
            this.btnDeclaracionRets.Name = "btnDeclaracionRets";
            this.btnDeclaracionRets.Size = new System.Drawing.Size(150, 43);
            this.btnDeclaracionRets.TabIndex = 7;
            this.btnDeclaracionRets.Text = "    RETS JSON";
            this.btnDeclaracionRets.UseVisualStyleBackColor = false;
            this.btnDeclaracionRets.Click += new System.EventHandler(this.btnDeclaracionRets_Click);
            // 
            // cbVentasElectronicas
            // 
            this.cbVentasElectronicas.AutoSize = true;
            this.cbVentasElectronicas.Location = new System.Drawing.Point(907, 254);
            this.cbVentasElectronicas.Name = "cbVentasElectronicas";
            this.cbVentasElectronicas.Size = new System.Drawing.Size(199, 29);
            this.cbVentasElectronicas.TabIndex = 10;
            this.cbVentasElectronicas.Text = "Ventas electrónicas";
            this.cbVentasElectronicas.UseVisualStyleBackColor = true;
            // 
            // FormBienvenida
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(32)))), ((int)(((byte)(40)))));
            this.ClientSize = new System.Drawing.Size(1440, 921);
            this.Controls.Add(this.btnDeclaracionRets);
            this.Controls.Add(this.btnExportDeclaracionJson);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Cliente);
            this.Controls.Add(this.btnExportar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FormBienvenida";
            this.Text = "FormBienvenida";
            this.Load += new System.EventHandler(this.FormBienvenida_Load);
            this.Cliente.ResumeLayout(false);
            this.Cliente.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.GroupBox Cliente;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbRepresentanteLegal;
        private System.Windows.Forms.TextBox tbDiaDeclaracion;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbMrl;
        private System.Windows.Forms.TextBox tbMt;
        private System.Windows.Forms.TextBox tbIess1;
        private System.Windows.Forms.TextBox tbSri;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox tbSuper;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox tbRazonSocial;
        private System.Windows.Forms.TextBox tbEstablecimiento;
        private System.Windows.Forms.TextBox tbContabilidad;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.TextBox tbIess2;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox tbRuc;
        private System.Windows.Forms.TextBox tbMes;
        private System.Windows.Forms.TextBox tbAnio;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnExportDeclaracionJson;
        private System.Windows.Forms.CheckBox cbDeclaracionSemestral;
        private System.Windows.Forms.Button btnDeclaracionRets;
        private System.Windows.Forms.CheckBox cbVentasElectronicas;
    }
}