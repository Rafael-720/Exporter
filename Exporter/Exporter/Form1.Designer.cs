namespace Exporter
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btExportar = new Button();
            pictureBox1 = new PictureBox();
            notifyIcon = new NotifyIcon(components);
            contextMenuStrip = new ContextMenuStrip(components);
            toolStripMenuItem1 = new ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            contextMenuStrip.SuspendLayout();
            SuspendLayout();
            // 
            // btExportar
            // 
            btExportar.BackColor = SystemColors.ActiveCaption;
            btExportar.Location = new Point(552, 254);
            btExportar.Name = "btExportar";
            btExportar.Size = new Size(123, 69);
            btExportar.TabIndex = 0;
            btExportar.Text = "Exportar";
            btExportar.UseVisualStyleBackColor = false;
            btExportar.Click += btExportar_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Image = Properties.Resources.SL_101820_36860_14;
            pictureBox1.Location = new Point(0, 1);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(697, 346);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 1;
            pictureBox1.TabStop = false;
            // 
            // notifyIcon
            // 
            notifyIcon.ContextMenuStrip = contextMenuStrip;
            notifyIcon.Icon = (Icon)resources.GetObject("notifyIcon.Icon");
            notifyIcon.Text = "notifyIcon1";
            notifyIcon.Visible = true;
            notifyIcon.MouseDoubleClick += notifyIcon_DoubleClick;
            // 
            // contextMenuStrip
            // 
            contextMenuStrip.ImageScalingSize = new Size(24, 24);
            contextMenuStrip.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1 });
            contextMenuStrip.Name = "contextMenuStrip";
            contextMenuStrip.Size = new Size(110, 26);
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(109, 22);
            toolStripMenuItem1.Text = "Fechar";
            toolStripMenuItem1.Click += fecharToolStripMenuItem_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(698, 342);
            Controls.Add(btExportar);
            Controls.Add(pictureBox1);
            Name = "Form1";
            Text = "Form1";
            Resize += Form1_Resize;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            contextMenuStrip.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button btExportar;
        private PictureBox pictureBox1;
        private NotifyIcon notifyIcon;
        private ContextMenuStrip contextMenuStrip;
        private ToolStripMenuItem toolStripMenuItem1;
    }
}