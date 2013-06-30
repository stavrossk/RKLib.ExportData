// ---------------------------------------------------------
// Rama Krishna's Export class
// Copyright (C) 2004 Rama Krishna. All rights reserved.
// ---------------------------------------------------------

using System;
using System.Data;
using RKLib.DatasetExporter.XslExporters;


namespace RKLib.DatasetExporter
{



	# region Summary

	/// <summary>
	/// Exports datatable to CSV or Excel format.
	/// This uses DataSet's XML features and XSLT for exporting.
	/// 
	/// C#.Net Example to be used in WebForms
	/// ------------------------------------- 
	/// using MyLib.ExportData;
	/// 
	/// private void btnExport_Click(object sender, System.EventArgs e)
	/// {
	///   try
	///   {
	///     // Declarations
	///     DataSet dsUsers =  ((DataSet) Session["dsUsers"]).Copy( );
	///     MyLib.ExportData.Export oExport = new MyLib.ExportData.Export("Web"); 
	///     string FileName = "UserList.csv";
	///     int[] ColList = {2, 3, 4, 5, 6};
	///     oExport.ExportDetails(dsUsers.Tables[0], ColList, Export.ExportFormat.CSV, FileName);
	///   }
	///   catch(Exception Ex)
	///   {
	///     lblError.Text = Ex.Message;
	///   }
	/// }	
	///  
	/// VB.Net Example to be used in WindowsForms
	/// ----------------------------------------- 
	/// Imports MyLib.ExportData
	/// 
	/// Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
	/// 
	///	  Try	
	///	  
	///     'Declarations
	/// 	Dim dsUsers As DataSet = (CType(Session("dsUsers"), DataSet)).Copy()
	/// 	Dim oExport As New MyLib.ExportData.Export("Win")
	/// 	Dim FileName As String = "C:\\UserList.xls"
	/// 	Dim ColList() As Integer = New Integer() {2, 3, 4, 5, 6}			
	///     oExport.ExportDetails(dsUsers.Tables(0), ColList, Export.ExportFormat.CSV, FileName)	 
	///     
	///   Catch Ex As Exception
	/// 	lblError.Text = Ex.Message
	///   End Try
	///   
	/// End Sub
	///     
	/// </summary>

	# endregion // Summary





	public class DatasetExporter
	{


        internal  enum ExportFormat
        { 
            Csv = 1,
            Excel = 2 
        }; // Export format enumeration


	    private readonly string _appType;	
			
		public DatasetExporter()
		{
			_appType = "Web";
		}




		public DatasetExporter(string applicationType)
		{

			_appType = applicationType;
			if(_appType != "Web" && _appType != "Win")
                throw new Exception
                    ("Provide valid application format (Web/Win)");
			
            if (_appType == "Web")
            {
            }


		}
		


		#region ExportDetails OverLoad : Type#1
		
		// Function  : ExportDetails 
		// Arguments : DetailsTable, FormatType, FileName
		// Purpose	 : To get all the column headers in the datatable and 
		//			   exorts in CSV / Excel format with all columns

		internal void ExportDetails
            (DataTable detailsTable, 
            ExportFormat FormatType, 
            string FileName)
		{

			try
			{	
			


				if(detailsTable.Rows.Count == 0) 
					throw new Exception
                        ("There are no details to export.");				
				

				// Create Dataset
				
                DataSet dsExport = new DataSet("Export");
				DataTable dtExport = detailsTable.Copy();
				
                dtExport.TableName = "Values"; 
				
                dsExport.Tables.Add(dtExport);	
				
				// Getting Field Names
				string[] sHeaders = new string[dtExport.Columns.Count];
				string[] sFileds = new string[dtExport.Columns.Count];
				
				for (int i=0; i < dtExport.Columns.Count; i++)
				{
					//sHeaders[i] = ReplaceSpclChars(dtExport.Columns[i].ColumnName);
					sHeaders[i] = dtExport.Columns[i].ColumnName;
					sFileds[i] = ReplaceSpecialCharacters(dtExport.Columns[i].ColumnName);					
				}

                if (_appType == "Web")
                {

                    WebXsltExporter.ExportWithXsltWeb
                        (dsExport, sHeaders, sFileds,
                        FormatType, FileName);

                }

                if (_appType == "Win")
                {

                    LocalXsltExporter.ExportWithXsltWindows
                        (dsExport, sHeaders, sFileds,
                        FormatType, FileName);
                
                }
            
            }			
			catch(Exception Ex)
			{
				throw Ex;
			}			
		}

		#endregion // ExportDetails OverLoad : Type#1

		#region ExportDetails OverLoad : Type#2

		// Function  : ExportDetails 
		// Arguments : DetailsTable, ColumnList, FormatType, FileName		
		// Purpose	 : To get the specified column headers in the datatable and
		//			   exorts in CSV / Excel format with specified columns


		internal void ExportDetails
            (DataTable detailsTable,
            int[] columnList, ExportFormat FormatType,
            string FileName)
		{


			try
			{


				if(detailsTable.Rows.Count == 0)
					throw new Exception
                        ("There are no details to export");
				
				// Create Dataset
				DataSet dsExport = new DataSet("Export");
				
                DataTable dtExport = detailsTable.Copy();
				
                dtExport.TableName = "Values"; 
				
                dsExport.Tables.Add(dtExport);

				if(columnList.Length > dtExport.Columns.Count)
					throw new Exception
                        ("ExportColumn List should not exceed Total Columns");
				

				// Getting Field Names
				var sHeaders = new string[columnList.Length];
				var sFileds = new string[columnList.Length];
				

				for (int i=0; i < columnList.Length; i++)
				{

					if((columnList[i] < 0) 
                        || (columnList[i] >= dtExport.Columns.Count))
						throw new Exception
                            ("ExportColumn Number should not exceed Total Columns Range");
					
					sHeaders[i] = dtExport.Columns[columnList[i]].ColumnName;
					sFileds[i] = ReplaceSpecialCharacters(dtExport.Columns[columnList[i]].ColumnName);					
				
                }


                switch (_appType)
                {
                    case "Web":
                        WebXsltExporter.ExportWithXsltWeb
                            (dsExport, sHeaders, sFileds,
                             FormatType, FileName);
                        break;
                    case "Win":
                        LocalXsltExporter.ExportWithXsltWindows
                            (dsExport, sHeaders, sFileds,
                             FormatType, FileName);
                        break;
                }


            }			
			catch(Exception Ex)
			{
				throw Ex;
			}			
		}
		
		#endregion // ExportDetails OverLoad : Type#2






		#region ExportDetails OverLoad : Type#3

		// Function  : ExportDetails 
		// Arguments : DetailsTable, ColumnList, Headers, FormatType, FileName	
		// Purpose	 : To get the specified column headers in the datatable and	
		//			   exorts in CSV / Excel format with specified columns and 
		//			   with specified headers

        internal void ExportDetails
            (DataTable detailsTable, int[] columnList,
            string[] headers, ExportFormat FormatType,
            string FileName)
        {
            if (detailsTable.Rows.Count == 0)
                throw new Exception("There are no details to export");

            // Create Dataset
            var dsExport = new DataSet("Export");
            
            DataTable dtExport = detailsTable.Copy();
                
            dtExport.TableName = "Values";
                
            dsExport.Tables.Add(dtExport);


            if (columnList.Length != headers.Length)
                throw new Exception
                    ("ExportColumn List and Headers List should be of same length");

            if (columnList.Length > dtExport.Columns.Count || headers.Length > dtExport.Columns.Count)
                throw new Exception
                    ("ExportColumn List should not exceed Total Columns");


            // Getting Field Names
            string[] sFileds = new string[columnList.Length];



            for (int i = 0; i < columnList.Length; i++)
            {

                if ((columnList[i] < 0) || (columnList[i] >= dtExport.Columns.Count))
                    throw new Exception
                        ("ExportColumn Number should not exceed Total Columns Range");

                sFileds[i] = ReplaceSpecialCharacters
                    (dtExport.Columns[columnList[i]].ColumnName);


            }


            PerformXsltExport
                (headers, FormatType,
                 FileName, dsExport, sFileds);
        }





        private void PerformXsltExport
            (string[] headers, ExportFormat formatType,
            string fileName, DataSet dsExport, string[] sFileds)
        {

            if (_appType == "Web")
                WebXsltExporter.ExportWithXsltWeb
                    (dsExport, headers, sFileds, formatType, fileName);

            if (_appType == "Win")
                LocalXsltExporter.ExportWithXsltWindows
                    (dsExport, headers, sFileds, formatType, fileName);
        }

		#endregion // ExportDetails OverLoad : Type#3






	    // Function  : ReplaceSpclChars 
		// Arguments : fieldName
		// Purpose   : Replaces special characters with XML codes 
		private static string ReplaceSpecialCharacters(string fieldName)
		{
			//			space 	-> 	_x0020_
			//			%		-> 	_x0025_
			//			#		->	_x0023_
			//			&		->	_x0026_
			//			/		->	_x002F_

			fieldName = fieldName.Replace(" ", "_x0020_");
			fieldName = fieldName.Replace("%", "_x0025_");
			fieldName = fieldName.Replace("#", "_x0023_");
			fieldName = fieldName.Replace("&", "_x0026_");
			fieldName = fieldName.Replace("/", "_x002F_");
			return fieldName;
		}




	}



    
}
