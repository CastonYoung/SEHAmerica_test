namespace SEHAmerica_ppt_Maker
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.label1 = new System.Windows.Forms.Label();
			this.TitleBox = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.BodyTextBox = new System.Windows.Forms.RichTextBox();
			this.ImageBtn = new System.Windows.Forms.Button();
			this.WriteBtn = new System.Windows.Forms.Button();
			this.Read_Btn = new System.Windows.Forms.Button();
			this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
			this.ListView1 = new System.Windows.Forms.ListView();
			this.Menu = new System.Windows.Forms.MenuStrip();
			this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.CheckNativeBoxes = new System.Windows.Forms.ToolStripMenuItem();
			this.Bold_Btn = new System.Windows.Forms.Button();
			this.Menu.SuspendLayout();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(33, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(55, 25);
			this.label1.TabIndex = 0;
			this.label1.Text = "Title:";
			// 
			// TitleBox
			// 
			this.TitleBox.AccessibleName = "TitleBox";
			this.TitleBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TitleBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.TitleBox.Location = new System.Drawing.Point(95, 44);
			this.TitleBox.MinimumSize = new System.Drawing.Size(152, 26);
			this.TitleBox.Name = "TitleBox";
			this.TitleBox.Size = new System.Drawing.Size(197, 26);
			this.TitleBox.TabIndex = 1;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(25, 94);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(63, 25);
			this.label2.TabIndex = 2;
			this.label2.Text = "Body:";
			// 
			// BodyTextBox
			// 
			this.BodyTextBox.AcceptsTab = true;
			this.BodyTextBox.AccessibleName = "BodyTextBox";
			this.BodyTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.BodyTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.BodyTextBox.ImeMode = System.Windows.Forms.ImeMode.On;
			this.BodyTextBox.Location = new System.Drawing.Point(95, 94);
			this.BodyTextBox.Name = "BodyTextBox";
			this.BodyTextBox.Size = new System.Drawing.Size(538, 157);
			this.BodyTextBox.TabIndex = 3;
			this.BodyTextBox.Text = "";
			this.BodyTextBox.SelectionChanged += new System.EventHandler(this.BodyTextBox_SelectionChanged);
			this.BodyTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.B);
			// 
			// ImageBtn
			// 
			this.ImageBtn.AccessibleName = "ImageBtn";
			this.ImageBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ImageBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ImageBtn.Location = new System.Drawing.Point(639, 192);
			this.ImageBtn.Name = "ImageBtn";
			this.ImageBtn.Size = new System.Drawing.Size(81, 59);
			this.ImageBtn.TabIndex = 4;
			this.ImageBtn.Text = "Image Search";
			this.ImageBtn.UseVisualStyleBackColor = true;
			this.ImageBtn.Click += new System.EventHandler(this.SearchImages);
			// 
			// WriteBtn
			// 
			this.WriteBtn.AccessibleName = "SaveBtn";
			this.WriteBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.WriteBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.WriteBtn.Location = new System.Drawing.Point(639, 118);
			this.WriteBtn.Name = "WriteBtn";
			this.WriteBtn.Size = new System.Drawing.Size(81, 36);
			this.WriteBtn.TabIndex = 5;
			this.WriteBtn.Text = "New Presentation";
			this.WriteBtn.UseVisualStyleBackColor = true;
			this.WriteBtn.Click += new System.EventHandler(this.Save);
			// 
			// Read_Btn
			// 
			this.Read_Btn.AccessibleName = "ReadBtn";
			this.Read_Btn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.Read_Btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Read_Btn.Location = new System.Drawing.Point(639, 44);
			this.Read_Btn.Name = "Read_Btn";
			this.Read_Btn.Size = new System.Drawing.Size(81, 33);
			this.Read_Btn.TabIndex = 6;
			this.Read_Btn.Text = "Load";
			this.Read_Btn.UseVisualStyleBackColor = true;
			this.Read_Btn.Click += new System.EventHandler(this.ReadSlide);
			// 
			// ListView1
			// 
			this.ListView1.AccessibleName = "ListView1";
			this.ListView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.ListView1.HideSelection = false;
			this.ListView1.Location = new System.Drawing.Point(95, 257);
			this.ListView1.MinimumSize = new System.Drawing.Size(152, 16);
			this.ListView1.Name = "ListView1";
			this.ListView1.Size = new System.Drawing.Size(625, 174);
			this.ListView1.TabIndex = 13;
			this.ListView1.TileSize = new System.Drawing.Size(160, 160);
			this.ListView1.UseCompatibleStateImageBehavior = false;
			this.ListView1.ItemActivate += new System.EventHandler(this.LoadImages);
			// 
			// Menu
			// 
			this.Menu.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.Menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
			this.Menu.Location = new System.Drawing.Point(0, 0);
			this.Menu.Name = "Menu";
			this.Menu.Size = new System.Drawing.Size(800, 28);
			this.Menu.TabIndex = 15;
			this.Menu.Text = "Menu";
			// 
			// settingsToolStripMenuItem
			// 
			this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.CheckNativeBoxes});
			this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
			this.settingsToolStripMenuItem.Size = new System.Drawing.Size(74, 24);
			this.settingsToolStripMenuItem.Text = "Settings";
			// 
			// CheckNativeBoxes
			// 
			this.CheckNativeBoxes.Checked = true;
			this.CheckNativeBoxes.CheckOnClick = true;
			this.CheckNativeBoxes.CheckState = System.Windows.Forms.CheckState.Checked;
			this.CheckNativeBoxes.Name = "CheckNativeBoxes";
			this.CheckNativeBoxes.Size = new System.Drawing.Size(254, 26);
			this.CheckNativeBoxes.Text = "Use Internal Rich Text Box";
			this.CheckNativeBoxes.CheckedChanged += new System.EventHandler(this.UseNativeTextBoxes);
			// 
			// Bold_Btn
			// 
			this.Bold_Btn.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("Bold_Btn.BackgroundImage")));
			this.Bold_Btn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
			this.Bold_Btn.FlatAppearance.BorderSize = 0;
			this.Bold_Btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.Bold_Btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Bold_Btn.Location = new System.Drawing.Point(64, 130);
			this.Bold_Btn.Margin = new System.Windows.Forms.Padding(0);
			this.Bold_Btn.Name = "Bold_Btn";
			this.Bold_Btn.Padding = new System.Windows.Forms.Padding(1, 0, 0, 0);
			this.Bold_Btn.Size = new System.Drawing.Size(24, 24);
			this.Bold_Btn.TabIndex = 16;
			this.Bold_Btn.UseVisualStyleBackColor = true;
			this.Bold_Btn.Click += new System.EventHandler(this.BoldBtn_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 469);
			this.Controls.Add(this.Bold_Btn);
			this.Controls.Add(this.Menu);
			this.Controls.Add(this.ListView1);
			this.Controls.Add(this.Read_Btn);
			this.Controls.Add(this.WriteBtn);
			this.Controls.Add(this.ImageBtn);
			this.Controls.Add(this.BodyTextBox);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.TitleBox);
			this.Controls.Add(this.label1);
			this.MainMenuStrip = this.Menu;
			this.MinimumSize = new System.Drawing.Size(95, 300);
			this.Name = "Form1";
			this.Text = "Form1";
			this.Menu.ResumeLayout(false);
			this.Menu.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox TitleBox;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button ImageBtn;
		private System.Windows.Forms.Button WriteBtn;
		private System.Windows.Forms.Button Read_Btn;
		private System.ComponentModel.BackgroundWorker backgroundWorker1;
		private System.Windows.Forms.ListView ListView1;
		public System.Windows.Forms.RichTextBox BodyTextBox;
		private System.Windows.Forms.MenuStrip Menu;
		private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem CheckNativeBoxes;
		private System.Windows.Forms.Button Bold_Btn;
	}
}

