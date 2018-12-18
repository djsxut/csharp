/*
 * 用户：Jason
 * 日期: 2018/12/18
 * 时间: 22:28
 */
using System;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

namespace demo
{
	/// <summary>
	/// Description of OutlookHelper.
	/// </summary>
	public class OutlookHelper
	{
	    Outlook.Application outlookApp_;
	    Outlook.NameSpace outlookNS_;
	    Outlook.MAPIFolder inbox_, sentBox_;
	    List<string> subject_ = new List<string>();
        
		public OutlookHelper()
		{
			this.LoadMail();
		}
		
		public int GetTotalMailCount()
		{
			return subject_.Count;
		}
		
		private void Open()
		{
			if (outlookApp_ == null)
			{
				outlookApp_ = new Outlook.Application();
	            outlookNS_ = outlookApp_.GetNamespace("MAPI");
	            inbox_ = outlookNS_.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
	            //sentbox_ = OutlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
			}
		}

		void LoadMail()
		{
			try
            {
                this.Open();
                for (int InboxIndx = 1; InboxIndx <= inbox_.Items.Count; InboxIndx++)
                {
                    try
                    {
                        var item = inbox_.Items[InboxIndx];
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mailItem = item as Outlook.MailItem;
                            
                            // demo只使用subject
                            if (!string.IsNullOrEmpty(mailItem.Subject))
                            {
                            	subject_.Add(mailItem.Subject);
                            }
                        }
                    }
                    catch (Exception ex1)
                    {
                        
                    }
                }
			}
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
            	outlookApp_ = null;
            }
		}
		
        public List<string> SearchSubject(string subject)
        {
        	List<string> result = new List<string>();
            try
            {
                foreach (string subject_item in subject_)
                {
	                if (subject_item.Contains(subject))
	                {
	                	result.Add(subject_item);
	                }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
            return result;
        }
	}
}
