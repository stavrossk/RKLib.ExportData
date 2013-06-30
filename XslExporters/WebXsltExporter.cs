using System;
using System.IO;
using System.Text;
using System.Data;
using System.Web;
using System.Xml;
using System.Xml.Xsl;
using System.Threading;

namespace RKLib.DatasetExporter.XslExporters
{


    internal static class WebXsltExporter
    {





        // Function  : Export_with_XSLT_Web 
        // Arguments : dsExport, sHeaders, sFileds, FormatType, FileName
        // Purpose   : Exports dataset into CSV / Excel format
        internal static void ExportWithXsltWeb
            (DataSet datasetToExport, string[] sHeaders,
            string[] sFileds, DatasetExporter.ExportFormat formatType, string fileName)
        {


            try
            {

                TextWriter textWriter = new StringWriter();
                
                var response = new HttpResponse(textWriter);



                ConfigureHttpResponse
                    (formatType, fileName, response);



                // XSLT to use for transforming this dataset.						
                var stream 
                    = new MemoryStream();

                var writer 
                    = new XmlTextWriter
                        (stream, Encoding.UTF8);


                XslStylesheetConstructor.CreateStylesheet
                    (writer, sHeaders, sFileds, formatType);
                
                
                writer.Flush();


                               
                var stringWriter
                    = PerformXslTransform
                    (datasetToExport, stream);


                PerformResponseWrite
                    (writer, stream,
                    response, stringWriter);
           
            
            }
            catch (ThreadAbortException Ex)
            {
                string errMsg = Ex.Message;
            }


            catch (Exception Ex)
            {
                throw Ex;
            }


        }



        private static StringWriter PerformXslTransform
            (DataSet dsExport, Stream stream)
        {


            stream.Seek(0, SeekOrigin.Begin);

            var xmlDoc = new XmlDataDocument(dsExport);

            //dsExport.WriteXml("Data.xml");
            XslTransform xslTransform = new XslTransform();


            xslTransform.Load
                (new XmlTextReader(stream),
                 null, null);


            var sw = new StringWriter();


            xslTransform.Transform
                (xmlDoc, null, sw, null);
            
            
            return sw;
        }




        private static void PerformResponseWrite
            (XmlWriter writer, Stream stream,
            HttpResponse response, TextWriter sw)
        {


            //xslTran.Transform
            //(System.Web.HttpContext.Current.Server.MapPath("Data.xml"),
            //null, sw, null);

            //Writeout the Content				
            response.Write
                (sw.ToString());

            sw.Close();

            writer.Close();

            stream.Close();

            response.End();
        
        }





        private static void ConfigureHttpResponse
            (DatasetExporter.ExportFormat formatType,
            string fileName, HttpResponse response)
        {


            // Appending Headers
            response.Clear();
            response.Buffer = true;

            if (formatType == DatasetExporter.ExportFormat.Csv)
            {

                response.ContentType 
                    = "text/csv";
                
                response.AppendHeader
                    ("content-disposition",
                    "attachment; filename=" + fileName);
            
            }
            else
            {

                response.ContentType 
                    = "application/vnd.ms-excel";

                response.AppendHeader
                    ("content-disposition",
                     "attachment; filename=" + fileName);
            
            }


        }





    }



}
