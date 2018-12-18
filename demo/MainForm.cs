/*
 * 用户：Jason
 * 日期: 2018/12/18
 * 时间: 1:11
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace demo
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		OutlookHelper outlook_;
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			// 启动时构造获取邮箱收件箱信息
			outlook_ = new OutlookHelper();
			this.labelMailInfo.Text = string.Format("共{0}封邮件", outlook_.GetTotalMailCount());
		}
		
        bool CheckExcelRowColValid()
        {
            try 
            {
                int row = Convert.ToInt32(this.textBoxExcelRow.Text);
                if (0 >= row) 
                {
                    throw new Exception("row of excel must a positive integer");;
                }
            } 
            catch (Exception) 
            {
                return false;
            }
                    
            string pattern=@"^[A-Za-z]+$";
            Regex  regex = new Regex(pattern);
            return regex.IsMatch(this.textBoxExcelCol.Text);
        }
        
		void ButtonExcelWriteClick(object sender, EventArgs e)
		{
            if (!CheckExcelRowColValid()) 
            {
                MessageBox.Show("please input excel row(1,2,3...) and col(A,B,C...)");
                return;
            }
			
			ExcelHelper excel = new ExcelHelper(this.textBoxExcelPath.Text);
			excel.SetValue(this.textBoxExcelRow.Text,
			               this.textBoxExcelCol.Text,
			               this.textBoxExcelValue.Text);
			excel.Close();
			excel = null;
		}
		
		void ButtonExcelReadClick(object sender, EventArgs e)
		{
            if (!CheckExcelRowColValid()) 
            {
                MessageBox.Show("please input excel row(1,2,3...) and col(A,B,C...)");
                return;
            }
			
			ExcelHelper excel = new ExcelHelper(this.textBoxExcelPath.Text);
			this.textBoxExcelValue.Text = excel.GetValue(this.textBoxExcelRow.Text,
			               								 this.textBoxExcelCol.Text);
			excel.Close();
			excel = null;
		}
		
		void LabelExcelSelectClick(object sender, EventArgs e)
		{
	        string excelPath = "";
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "excel Files|*.xls;*.xlsx";
            if (op.ShowDialog() == DialogResult.OK)
            {
                excelPath = op.FileName;
            }
            else
            {
                excelPath = "";
            }
            
            this.textBoxExcelPath.Text = excelPath;
		}
		
		void ButtonOpenExcelClick(object sender, EventArgs e)
		{
	        try 
	        {
                System.Diagnostics.Process.Start("excel");
            } 
	        catch (Exception) 
	        {
                MessageBox.Show("Excel not found, start failed!", 
	        	                "start Excel failed", 
	        	                MessageBoxButtons.OK, 
	        	                MessageBoxIcon.Error);
            }     
		}
		
		void ButtonOpenOutlookClick(object sender, EventArgs e)
		{
	        try 
	        {
                System.Diagnostics.Process.Start("Outlook");
            } 
	        catch (Exception)
	        {
                MessageBox.Show("Outlook not found, start failed!", 
	        	                "start Outlook failed", 
	        	                MessageBoxButtons.OK, 
	        	                MessageBoxIcon.Error);
            }
		}
		
		void ButtonMailSearchClick(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(this.textBoxMailSubject.Text))
			{
				return;
			}
			
			List<string> matchedSubject = outlook_.SearchSubject(this.textBoxMailSubject.Text);
			this.labelMailInfo.Text = string.Format("找到{0}/{1}封邮件", matchedSubject.Count, outlook_.GetTotalMailCount());
			
			string msg = "";
			int idx = 1;
			foreach (string subject in matchedSubject)
			{
				msg = msg + idx.ToString() + ": " + subject + "\n";
				idx += 1;
			}
			
            MessageBox.Show(msg, 
        	                "mail info", 
        	                MessageBoxButtons.OK, 
        	                MessageBoxIcon.Information);
		}
	}
}
