/*
 * 用户：Jason 
 * 日期: 2018/12/17
 * 时间: 21:25
 */
using System;
using System.IO;
using System.Data;
using System.Configuration;
using System.Web;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace demo
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class ExcelHelper
	{
		private string excelPath_;
		private bool opend_ = false;
		private Microsoft.Office.Interop.Excel.Application excel_;
		private Microsoft.Office.Interop.Excel.Workbooks wbs_;
        private Microsoft.Office.Interop.Excel.Workbook wb_;
        private object typeMissing_ = System.Reflection.Missing.Value;
        
		public ExcelHelper(string excelPath)
		{
			excelPath_ = excelPath;
			Open();
		}
		
		public bool Open()
		{
			excel_ = new Microsoft.Office.Interop.Excel.Application();
			wbs_ = excel_.Workbooks;
			
			// 文件存在就打开, 否则创建
            if (File.Exists(excelPath_))
            {
            	//wb_ = wbs_.Add(excelPath_);
            	wb_ = wbs_.Open(excelPath_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_, typeMissing_);
            }
            else
            {
            	wb_ = wbs_.Add(true);
            	return SaveAs();
            }
            
			return true;
		}
		
		// demo程序, 默认操作sheet-1
		private Microsoft.Office.Interop.Excel.Worksheet GetSheet()
		{
			 //return (Microsoft.Office.Interop.Excel.Worksheet)wb_.Sheets.get_Item(1);
			 return (Microsoft.Office.Interop.Excel.Worksheet)wb_.Sheets.get_Item(wb_.Sheets.Count);
		}
		
		public string GetValue(string row, string col)
		{
			Microsoft.Office.Interop.Excel.Worksheet ws = GetSheet();
			return ((Microsoft.Office.Interop.Excel.Range)ws.Cells[row, col]).Text.ToString();
		}
		
		public void SetValue(string row, string col, string value)
		{
			Microsoft.Office.Interop.Excel.Worksheet ws = GetSheet();
			ws.Cells[row, col] = value;
			Save();
		}
		
		public bool Save()
        {
            try
            {
                wb_.Save();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
		
        private bool SaveAs()
        {
            try
            {
                wb_.SaveAs(excelPath_, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true; 
            } 
            catch (Exception ex)
            {
                return false; 
            }
        }
        
        public void Close()
        {
        	try {
	        	wb_.Close(Type.Missing, Type.Missing, Type.Missing);
	            wbs_.Close();
	            excel_.Quit();
	            wb_ = null;
	            wbs_ = null;
	            excel_ = null;
	            GC.Collect();
        	} catch (Exception) {

        	}
        }
	}
}
