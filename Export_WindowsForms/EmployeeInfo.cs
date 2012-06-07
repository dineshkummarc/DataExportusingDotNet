using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using RKLib.ExportData;

namespace ExportDataGrid
{
	/// <summary>
	/// Summary description for EmployeeInfo.
	/// </summary>
	public class EmployeeInfo : System.Windows.Forms.Form
	{
		private DataSet dsEmployee = new DataSet();
		private System.Windows.Forms.DataGrid dgEmployee;
		private System.Windows.Forms.Button btnExportCSV;
		private System.Windows.Forms.Button btnExportExcel;
		private System.Windows.Forms.Label lblHeader;
		private System.Windows.Forms.Label lblMessage;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public EmployeeInfo()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.dgEmployee = new System.Windows.Forms.DataGrid();
			this.btnExportCSV = new System.Windows.Forms.Button();
			this.btnExportExcel = new System.Windows.Forms.Button();
			this.lblMessage = new System.Windows.Forms.Label();
			this.lblHeader = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dgEmployee)).BeginInit();
			this.SuspendLayout();
			// 
			// dgEmployee
			// 
			this.dgEmployee.BackgroundColor = System.Drawing.Color.White;
			this.dgEmployee.CaptionBackColor = System.Drawing.Color.MidnightBlue;
			this.dgEmployee.DataMember = "";
			this.dgEmployee.Font = new System.Drawing.Font("Garamond", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgEmployee.HeaderForeColor = System.Drawing.Color.White;
			this.dgEmployee.Location = new System.Drawing.Point(28, 77);
			this.dgEmployee.Name = "dgEmployee";
			this.dgEmployee.Size = new System.Drawing.Size(564, 307);
			this.dgEmployee.TabIndex = 0;
			// 
			// btnExportCSV
			// 
			this.btnExportCSV.BackColor = System.Drawing.Color.DarkGray;
			this.btnExportCSV.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExportCSV.Location = new System.Drawing.Point(40, 424);
			this.btnExportCSV.Name = "btnExportCSV";
			this.btnExportCSV.Size = new System.Drawing.Size(112, 23);
			this.btnExportCSV.TabIndex = 1;
			this.btnExportCSV.Text = "Export to CSV";
			this.btnExportCSV.Click += new System.EventHandler(this.btnExportCSV_Click);
			// 
			// btnExportExcel
			// 
			this.btnExportExcel.BackColor = System.Drawing.Color.DarkGray;
			this.btnExportExcel.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExportExcel.Location = new System.Drawing.Point(160, 424);
			this.btnExportExcel.Name = "btnExportExcel";
			this.btnExportExcel.Size = new System.Drawing.Size(128, 24);
			this.btnExportExcel.TabIndex = 3;
			this.btnExportExcel.Text = "Export to Excel";
			this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);
			// 
			// lblMessage
			// 
			this.lblMessage.Font = new System.Drawing.Font("Garamond", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMessage.ForeColor = System.Drawing.Color.Red;
			this.lblMessage.Location = new System.Drawing.Point(40, 472);
			this.lblMessage.Name = "lblMessage";
			this.lblMessage.Size = new System.Drawing.Size(544, 40);
			this.lblMessage.TabIndex = 4;
			this.lblMessage.Text = "Error Message";
			// 
			// lblHeader
			// 
			this.lblHeader.Font = new System.Drawing.Font("Garamond", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblHeader.Location = new System.Drawing.Point(32, 16);
			this.lblHeader.Name = "lblHeader";
			this.lblHeader.Size = new System.Drawing.Size(352, 48);
			this.lblHeader.TabIndex = 5;
			this.lblHeader.Text = "Exporting DataGrid into CSV / Excel File Author : Rama Krishna";
			// 
			// EmployeeInfo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.BackColor = System.Drawing.Color.WhiteSmoke;
			this.ClientSize = new System.Drawing.Size(648, 533);
			this.Controls.Add(this.lblHeader);
			this.Controls.Add(this.lblMessage);
			this.Controls.Add(this.btnExportExcel);
			this.Controls.Add(this.btnExportCSV);
			this.Controls.Add(this.dgEmployee);
			this.Font = new System.Drawing.Font("Garamond", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "EmployeeInfo";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "EmployeeInfo - WindowsForms C#";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.EmployeeInfo_Load);
			((System.ComponentModel.ISupportInitialize)(this.dgEmployee)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new EmployeeInfo());
		}


		private void EmployeeInfo_Load(object sender, System.EventArgs e)
		{
			dgEmployee.DataSource = get_AllEmployees().Tables["Employee"];
			lblMessage.Text = "";
		}

		private DataSet get_AllEmployees()
		{
			try
			{
				// Get the employee details
				string strSql = "SELECT EmployeeID, LastName, FirstName, Format(BirthDate, 'dd-MMM-yyyy') As BirthDate, Format(HireDate, 'dd-MMM-yyyy') As HireDate, Address, PostalCode FROM Employees";
				OleDbConnection objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath.Replace("bin\\Debug","") + "Database\\Employee.mdb");
				OleDbDataAdapter daEmp = new OleDbDataAdapter(strSql, objConn);				
				daEmp.Fill(dsEmployee, "Employee");
				return dsEmployee;
			}
			catch(Exception Ex)
			{
				throw Ex;
			}
		}

		private void btnExportCSV_Click(object sender, System.EventArgs e)
		{
			// Export all the details
			try
			{			
				lblMessage.Text = "";
				// Get the datatable to export			
				DataTable dtEmployee = dsEmployee.Tables["Employee"].Copy();
				
				// Specify the column list to export
				int[] iColumns = {1,2,3,5,6};
				
				// Export the details of specified columns to CSV
				RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Win");
				objExport.ExportDetails(dtEmployee, iColumns, Export.ExportFormat.CSV, "C:\\EmployeesInfo.csv");
				lblMessage.Text = "Successfully exported to C:\\EmployeesInfo.csv";				
			}
			catch(Exception Ex)
			{
				lblMessage.Text = Ex.Message;
			}
		}

		private void btnExportExcel_Click(object sender, System.EventArgs e)
		{
			lblMessage.Text = "";
			// Export all the details
			try
			{			
				// Get the datatable to export			
				DataTable dtEmployee = dsEmployee.Tables["Employee"].Copy();

				// Export all the details to Excel
				RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Win");				
				objExport.ExportDetails(dtEmployee, Export.ExportFormat.Excel, "C:\\EmployeesInfo.xls");
				lblMessage.Text = "Successfully exported to C:\\EmployeesInfo.xls";
			}
			catch(Exception Ex)
			{
				lblMessage.Text = Ex.Message;
			}
		}
	}
}
