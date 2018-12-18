/*
 * 由SharpDevelop创建。
 * 用户： HP
 * 日期: 2018/12/18
 * 时间: 1:11
 * 
 * 要改变这种模板请点击 工具|选项|代码编写|编辑标准头文件
 */
namespace demo
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TextBox textBoxExcelValue;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBoxExcelCol;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBoxExcelRow;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button buttonExcelWrite;
		private System.Windows.Forms.TextBox textBoxExcelPath;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label labelExcelSelect;
		private System.Windows.Forms.Button buttonExcelRead;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button buttonOpenOutlook;
		private System.Windows.Forms.Button buttonOpenExcel;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Button buttonMailSearch;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label labelMailInfo;
		private System.Windows.Forms.TextBox textBoxMailSubject;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.labelExcelSelect = new System.Windows.Forms.Label();
			this.buttonExcelRead = new System.Windows.Forms.Button();
			this.buttonExcelWrite = new System.Windows.Forms.Button();
			this.textBoxExcelPath = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.textBoxExcelValue = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.textBoxExcelCol = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.textBoxExcelRow = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.buttonOpenOutlook = new System.Windows.Forms.Button();
			this.buttonOpenExcel = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.labelMailInfo = new System.Windows.Forms.Label();
			this.buttonMailSearch = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.textBoxMailSubject = new System.Windows.Forms.TextBox();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.labelExcelSelect);
			this.groupBox1.Controls.Add(this.buttonExcelRead);
			this.groupBox1.Controls.Add(this.buttonExcelWrite);
			this.groupBox1.Controls.Add(this.textBoxExcelPath);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.textBoxExcelValue);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.textBoxExcelCol);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.textBoxExcelRow);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(28, 126);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(604, 175);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Excel操作";
			// 
			// labelExcelSelect
			// 
			this.labelExcelSelect.Location = new System.Drawing.Point(364, 140);
			this.labelExcelSelect.Name = "labelExcelSelect";
			this.labelExcelSelect.Size = new System.Drawing.Size(44, 24);
			this.labelExcelSelect.TabIndex = 10;
			this.labelExcelSelect.Text = "选择";
			this.labelExcelSelect.Click += new System.EventHandler(this.LabelExcelSelectClick);
			// 
			// buttonExcelRead
			// 
			this.buttonExcelRead.Location = new System.Drawing.Point(462, 128);
			this.buttonExcelRead.Name = "buttonExcelRead";
			this.buttonExcelRead.Size = new System.Drawing.Size(75, 34);
			this.buttonExcelRead.TabIndex = 9;
			this.buttonExcelRead.Text = "读取";
			this.buttonExcelRead.UseVisualStyleBackColor = true;
			this.buttonExcelRead.Click += new System.EventHandler(this.ButtonExcelReadClick);
			// 
			// buttonExcelWrite
			// 
			this.buttonExcelWrite.Location = new System.Drawing.Point(462, 78);
			this.buttonExcelWrite.Name = "buttonExcelWrite";
			this.buttonExcelWrite.Size = new System.Drawing.Size(75, 36);
			this.buttonExcelWrite.TabIndex = 8;
			this.buttonExcelWrite.Text = "设置";
			this.buttonExcelWrite.UseVisualStyleBackColor = true;
			this.buttonExcelWrite.Click += new System.EventHandler(this.ButtonExcelWriteClick);
			// 
			// textBoxExcelPath
			// 
			this.textBoxExcelPath.Location = new System.Drawing.Point(78, 133);
			this.textBoxExcelPath.Name = "textBoxExcelPath";
			this.textBoxExcelPath.Size = new System.Drawing.Size(281, 28);
			this.textBoxExcelPath.TabIndex = 7;
			this.textBoxExcelPath.Text = "D:\\demo.xlsx";
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("宋体", 12F);
			this.label4.Location = new System.Drawing.Point(5, 136);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(67, 28);
			this.label4.TabIndex = 6;
			this.label4.Text = "路径";
			// 
			// textBoxExcelValue
			// 
			this.textBoxExcelValue.Location = new System.Drawing.Point(78, 84);
			this.textBoxExcelValue.Name = "textBoxExcelValue";
			this.textBoxExcelValue.Size = new System.Drawing.Size(281, 28);
			this.textBoxExcelValue.TabIndex = 5;
			this.textBoxExcelValue.Text = "demo";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("宋体", 12F);
			this.label3.Location = new System.Drawing.Point(31, 84);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(53, 28);
			this.label3.TabIndex = 4;
			this.label3.Text = "值";
			// 
			// textBoxExcelCol
			// 
			this.textBoxExcelCol.Location = new System.Drawing.Point(295, 33);
			this.textBoxExcelCol.Name = "textBoxExcelCol";
			this.textBoxExcelCol.Size = new System.Drawing.Size(64, 28);
			this.textBoxExcelCol.TabIndex = 3;
			this.textBoxExcelCol.Text = "A";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("宋体", 12F);
			this.label2.Location = new System.Drawing.Point(236, 33);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 28);
			this.label2.TabIndex = 2;
			this.label2.Text = "列";
			// 
			// textBoxExcelRow
			// 
			this.textBoxExcelRow.Location = new System.Drawing.Point(78, 33);
			this.textBoxExcelRow.Name = "textBoxExcelRow";
			this.textBoxExcelRow.Size = new System.Drawing.Size(64, 28);
			this.textBoxExcelRow.TabIndex = 1;
			this.textBoxExcelRow.Text = "1";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("宋体", 12F);
			this.label1.Location = new System.Drawing.Point(31, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 28);
			this.label1.TabIndex = 0;
			this.label1.Text = "行";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.buttonOpenOutlook);
			this.groupBox2.Controls.Add(this.buttonOpenExcel);
			this.groupBox2.Location = new System.Drawing.Point(28, 29);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(604, 82);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "打开应用程序";
			// 
			// buttonOpenOutlook
			// 
			this.buttonOpenOutlook.Location = new System.Drawing.Point(317, 27);
			this.buttonOpenOutlook.Name = "buttonOpenOutlook";
			this.buttonOpenOutlook.Size = new System.Drawing.Size(136, 38);
			this.buttonOpenOutlook.TabIndex = 1;
			this.buttonOpenOutlook.Text = "打开Outlook";
			this.buttonOpenOutlook.UseVisualStyleBackColor = true;
			this.buttonOpenOutlook.Click += new System.EventHandler(this.ButtonOpenOutlookClick);
			// 
			// buttonOpenExcel
			// 
			this.buttonOpenExcel.Location = new System.Drawing.Point(116, 27);
			this.buttonOpenExcel.Name = "buttonOpenExcel";
			this.buttonOpenExcel.Size = new System.Drawing.Size(133, 38);
			this.buttonOpenExcel.TabIndex = 0;
			this.buttonOpenExcel.Text = "打开Excel";
			this.buttonOpenExcel.UseVisualStyleBackColor = true;
			this.buttonOpenExcel.Click += new System.EventHandler(this.ButtonOpenExcelClick);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.labelMailInfo);
			this.groupBox3.Controls.Add(this.buttonMailSearch);
			this.groupBox3.Controls.Add(this.label5);
			this.groupBox3.Controls.Add(this.textBoxMailSubject);
			this.groupBox3.Location = new System.Drawing.Point(33, 329);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(599, 100);
			this.groupBox3.TabIndex = 2;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Outlook操作(subject搜索)";
			// 
			// labelMailInfo
			// 
			this.labelMailInfo.Location = new System.Drawing.Point(381, 39);
			this.labelMailInfo.Name = "labelMailInfo";
			this.labelMailInfo.Size = new System.Drawing.Size(178, 28);
			this.labelMailInfo.TabIndex = 3;
			this.labelMailInfo.Text = "mail";
			// 
			// buttonMailSearch
			// 
			this.buttonMailSearch.Location = new System.Drawing.Point(295, 33);
			this.buttonMailSearch.Name = "buttonMailSearch";
			this.buttonMailSearch.Size = new System.Drawing.Size(75, 36);
			this.buttonMailSearch.TabIndex = 2;
			this.buttonMailSearch.Text = "搜索";
			this.buttonMailSearch.UseVisualStyleBackColor = true;
			this.buttonMailSearch.Click += new System.EventHandler(this.ButtonMailSearchClick);
			// 
			// label5
			// 
			this.label5.Font = new System.Drawing.Font("宋体", 12F);
			this.label5.Location = new System.Drawing.Point(26, 36);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(69, 36);
			this.label5.TabIndex = 1;
			this.label5.Text = "主题";
			// 
			// textBoxMailSubject
			// 
			this.textBoxMailSubject.Location = new System.Drawing.Point(99, 36);
			this.textBoxMailSubject.Name = "textBoxMailSubject";
			this.textBoxMailSubject.Size = new System.Drawing.Size(176, 28);
			this.textBoxMailSubject.TabIndex = 0;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(660, 483);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Name = "MainForm";
			this.Text = "demo";
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox3.PerformLayout();
			this.ResumeLayout(false);

		}
	}
}
