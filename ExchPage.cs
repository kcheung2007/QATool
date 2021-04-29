using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace QATool
{
	/// <summary>
	/// Summary description for Performance.
	/// </summary>
	public class ExchPage : System.Windows.Forms.UserControl
	{
        private ExchangeSDK.Tools.ExchangeTreeView.ExchangeTreeViewControl exchTreeViewControl1;
        private System.ComponentModel.IContainer components;

		public ExchPage()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.exchTreeViewControl1 = new ExchangeSDK.Tools.ExchangeTreeView.ExchangeTreeViewControl();
            this.SuspendLayout();
            // 
            // exchTreeViewControl1
            // 
            this.exchTreeViewControl1.Location = new System.Drawing.Point(32, 108);
            this.exchTreeViewControl1.Name = "exchTreeViewControl1";
            this.exchTreeViewControl1.Size = new System.Drawing.Size(308, 232);
            this.exchTreeViewControl1.TabIndex = 0;
            // 
            // ExchPage
            // 
            this.Controls.Add(this.exchTreeViewControl1);
            this.Name = "ExchPage";
            this.Size = new System.Drawing.Size(388, 428);
            this.ResumeLayout(false);

        }
		#endregion

	}
}
