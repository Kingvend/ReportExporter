namespace ReportExporter
{
	partial class Form1
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
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// simpleButton1
			// 
			this.simpleButton1.Dock = System.Windows.Forms.DockStyle.Top;
			this.simpleButton1.Location = new System.Drawing.Point(0, 20);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(495, 45);
			this.simpleButton1.TabIndex = 0;
			this.simpleButton1.Text = "Word export";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// textEdit1
			// 
			this.textEdit1.Dock = System.Windows.Forms.DockStyle.Top;
			this.textEdit1.Location = new System.Drawing.Point(0, 0);
			this.textEdit1.Name = "textEdit1";
			this.textEdit1.Size = new System.Drawing.Size(495, 20);
			this.textEdit1.TabIndex = 1;
			// 
			// simpleButton2
			// 
			this.simpleButton2.Dock = System.Windows.Forms.DockStyle.Top;
			this.simpleButton2.Location = new System.Drawing.Point(0, 65);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(495, 45);
			this.simpleButton2.TabIndex = 2;
			this.simpleButton2.Text = "Pdf export";
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// simpleButton3
			// 
			this.simpleButton3.Dock = System.Windows.Forms.DockStyle.Top;
			this.simpleButton3.Location = new System.Drawing.Point(0, 110);
			this.simpleButton3.Name = "simpleButton3";
			this.simpleButton3.Size = new System.Drawing.Size(495, 45);
			this.simpleButton3.TabIndex = 3;
			this.simpleButton3.Text = "Открыть файл";
			this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(495, 162);
			this.Controls.Add(this.simpleButton3);
			this.Controls.Add(this.simpleButton2);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.textEdit1);
			this.Name = "Form1";
			this.Text = "Form1";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.TextEdit textEdit1;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private DevExpress.XtraEditors.SimpleButton simpleButton3;
	}
}

