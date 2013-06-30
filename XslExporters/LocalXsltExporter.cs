using System;
using System.Text;
using System.Data;
using System.Xml;
using System.Xml.Xsl;

using System.IO;



namespace RKLib.DatasetExporter
{


    internal static class LocalXsltExporter
    {






        // Function  : Export_with_XSLT_Windows 
        // Arguments : dsExport, sHeaders, sFileds, FormatType, FileName
        // Purpose   : Exports dataset into CSV / Excel format
        internal static void ExportWithXsltWindows
            (DataSet datasetToExport, string[] sHeaders, string[] sFileds,
            DatasetExporter.ExportFormat FormatType, string FileName)
        {

            try
            {

                MemoryStream memoryStream;
                XmlTextWriter writer;


                ConstructXsltStylesheet
                    (sHeaders, sFileds, FormatType,
                    out memoryStream, out writer);


                StringWriter sw
                    = PerformXsltTransform
                    (datasetToExport, memoryStream);


                PerformFileWrite
                    (FileName, memoryStream, writer, sw);



            }
            catch (Exception Ex)
            {
                throw Ex;
            }



        }



        private static System.IO.StringWriter PerformXsltTransform
            (DataSet dsExport, MemoryStream memoryStream)
        {


            XmlDataDocument xmlDoc = new XmlDataDocument(dsExport);

            XslTransform xslTran = new XslTransform();

            xslTran.Load(new XmlTextReader(memoryStream), null, null);

            System.IO.StringWriter sw = new System.IO.StringWriter();

            xslTran.Transform(xmlDoc, null, sw, null);
            return sw;
        }


        private static void ConstructXsltStylesheet
            (string[] sHeaders, string[] sFileds,
            DatasetExporter.ExportFormat FormatType,
            out MemoryStream memoryStream,
            out XmlTextWriter writer)
        {


            // XSLT to use for transforming this dataset.						
            memoryStream = new MemoryStream();

            writer = new XmlTextWriter
                (memoryStream, Encoding.UTF8);

            XslStylesheetConstructor.CreateStylesheet
                (writer, sHeaders,
                sFileds, FormatType);

            writer.Flush();

            memoryStream.Seek
                (0, SeekOrigin.Begin);

        }


        private static void PerformFileWrite
            (string FileName, MemoryStream stream,
            XmlTextWriter writer, System.IO.StringWriter sw)
        {

            //Writeout the Content									
            StreamWriter strwriter = new StreamWriter(FileName);
            strwriter.WriteLine(sw.ToString());
            strwriter.Close();

            sw.Close();
            writer.Close();
            stream.Close();

        }









    }


}
