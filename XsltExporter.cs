using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.Xml.Xsl;
using RKLib.DatasetExporter;
using System.Threading;





namespace RKLib.DatasetExporter
{


    internal static class XsltExporter
    {





        // Function  : Export_with_XSLT_Web 
        // Arguments : dsExport, sHeaders, sFileds, FormatType, FileName
        // Purpose   : Exports dataset into CSV / Excel format
        internal static void Export_with_XSLT_Web
            (DataSet dsExport, string[] sHeaders,
            string[] sFileds, DatasetExporter.ExportFormat FormatType, string FileName)
        {


            try
            {

                System.Web.HttpResponse response = null;

                // Appending Headers
                //response.Clear();
                //response.Buffer = true;

                if (FormatType == DatasetExporter.ExportFormat.CSV)
                {
                    response.ContentType = "text/csv";
                    response.AppendHeader("content-disposition", "attachment; filename=" + FileName);
                }
                else
                {

                    response.ContentType = "application/vnd.ms-excel";
                
                    response.AppendHeader
                        ("content-disposition",
                        "attachment; filename=" + FileName);
                
                }

                // XSLT to use for transforming this dataset.						
                MemoryStream stream 
                    = new MemoryStream();

                XmlTextWriter writer 
                    = new XmlTextWriter
                        (stream, Encoding.UTF8);


                XslStylesheetConstructor.CreateStylesheet
                    (writer, sHeaders, sFileds, FormatType);
                
                
                writer.Flush();


                
                
               
                stream.Seek(0, SeekOrigin.Begin);

                XmlDataDocument xmlDoc = new XmlDataDocument(dsExport);
                //dsExport.WriteXml("Data.xml");
                XslTransform xslTran = new XslTransform();
                xslTran.Load(new XmlTextReader(stream), null, null);

                System.IO.StringWriter sw = new System.IO.StringWriter();
                xslTran.Transform(xmlDoc, null, sw, null);
                //xslTran.Transform(System.Web.HttpContext.Current.Server.MapPath("Data.xml"), null, sw, null);

                //Writeout the Content				
                response.Write(sw.ToString());
                sw.Close();
                writer.Close();
                stream.Close();
                response.End();
            }
            catch (ThreadAbortException Ex)
            {
                string ErrMsg = Ex.Message;
            }


            catch (Exception Ex)
            {
                throw Ex;
            }


        }





    }



}
