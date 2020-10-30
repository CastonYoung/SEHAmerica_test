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
			this.components = new System.ComponentModel.Container();
			this.label1 = new System.Windows.Forms.Label();
			this.TitleBox = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.BodyTextBox = new System.Windows.Forms.RichTextBox();
			this.ImageBtn = new System.Windows.Forms.Button();
			this.SaveBtn = new System.Windows.Forms.Button();
			this.ReadBtn = new System.Windows.Forms.Button();
			this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
			this.ListView1 = new System.Windows.Forms.ListView();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
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
			// SaveBtn
			// 
			this.SaveBtn.AccessibleName = "SaveBtn";
			this.SaveBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.SaveBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.SaveBtn.Location = new System.Drawing.Point(639, 118);
			this.SaveBtn.Name = "SaveBtn";
			this.SaveBtn.Size = new System.Drawing.Size(81, 36);
			this.SaveBtn.TabIndex = 5;
			this.SaveBtn.Text = "Save";
			this.SaveBtn.UseVisualStyleBackColor = true;
			this.SaveBtn.Click += new System.EventHandler(this.Save);
			// 
			// ReadBtn
			// 
			this.ReadBtn.AccessibleName = "ReadBtn";
			this.ReadBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ReadBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ReadBtn.Location = new System.Drawing.Point(639, 44);
			this.ReadBtn.Name = "ReadBtn";
			this.ReadBtn.Size = new System.Drawing.Size(81, 33);
			this.ReadBtn.TabIndex = 6;
			this.ReadBtn.Text = "Read";
			this.ReadBtn.UseVisualStyleBackColor = true;
			this.ReadBtn.Click += new System.EventHandler(this.ReadSlide);
			// 
			// ListView1
			// 
			this.ListView1.AccessibleName = "ListView1";
			this.ListView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.ListView1.HideSelection = false;
			this.ListView1.Location = new System.Drawing.Point(95, 257);
			this.ListView1.Name = "ListView1";
			this.ListView1.Size = new System.Drawing.Size(625, 174);
			this.ListView1.TabIndex = 13;
			this.ListView1.TileSize = new System.Drawing.Size(160, 160);
			this.ListView1.UseCompatibleStateImageBehavior = false;
			this.ListView1.ItemActivate += new System.EventHandler(this.NewLoadImages);
			// 
			// imageList1
			// 
			this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// toolStrip1
			// 
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Size = new System.Drawing.Size(800, 25);
			this.toolStrip1.TabIndex = 14;
			this.toolStrip1.Text = "toolStrip1";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 469);
			this.Controls.Add(this.toolStrip1);
			this.Controls.Add(this.ListView1);
			this.Controls.Add(this.ReadBtn);
			this.Controls.Add(this.SaveBtn);
			this.Controls.Add(this.ImageBtn);
			this.Controls.Add(this.BodyTextBox);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.TitleBox);
			this.Controls.Add(this.label1);
			this.Name = "Form1";
			this.Text = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox TitleBox;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.RichTextBox BodyTextBox;
		private System.Windows.Forms.Button ImageBtn;
		private System.Windows.Forms.Button SaveBtn;
		private System.Windows.Forms.Button ReadBtn;
		private System.ComponentModel.BackgroundWorker backgroundWorker1;
		private System.Windows.Forms.ListView ListView1;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolStrip toolStrip1;
	}
}

