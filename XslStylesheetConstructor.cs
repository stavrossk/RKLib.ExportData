using System.Collections.Generic;
using System.Xml;

namespace RKLib.DatasetExporter
{


    internal class XslStylesheetConstructor
    {



        // Function  : WriteStylesheet 
        // Arguments : writer, sHeaders, sFileds, FormatType
        // Purpose   : Creates XSLT file to apply on dataset's XML file 
        internal static void CreateStylesheet
            (XmlTextWriter writer, string[] sHeaders,
            string[] sFileds, DatasetExporter.ExportFormat formatType)
        {


            ConstructXslSkeleton(writer);

            // xsl-template
            writer.WriteStartElement("xsl:template");
            writer.WriteAttributeString("match", "/");


            XslValueOfForHeaders
                (writer, sHeaders,
                 sFileds, formatType);


            XslForEach(writer);


            XslValueOfForDataFields
                (writer, sFileds, formatType);


            FinalizeXslStylesheet(writer);

        
        }




        private static void FinalizeXslStylesheet(XmlWriter writer)
        {
            
            writer.WriteEndElement(); // xsl:for-each
            
            writer.WriteEndElement(); // xsl-template
            
            writer.WriteEndElement(); // xsl:stylesheet
        
            writer.WriteEndDocument();
        
        }


        private static void XslForEach(XmlWriter writer)
        {

            // xsl:for-each
            writer.WriteStartElement
                ("xsl:for-each");
            
            writer.WriteAttributeString
                ("select", "Export/Values");
            
            writer.WriteString
                ("\r\n");
        
        }



        private static void XslValueOfForDataFields
            (XmlWriter writer, IList<string> sFileds,
            DatasetExporter.ExportFormat formatType)
        {

            // xsl:value-of for data fields
            for (int i = 0; i < sFileds.Count; i++)
            {
                writer.WriteString("\"");
                writer.WriteStartElement("xsl:value-of");
                writer.WriteAttributeString("select", sFileds[i]);
                writer.WriteEndElement(); // xsl:value-of
                writer.WriteString("\"");

                if (i != sFileds.Count - 1)
                {
                    writer.WriteString
                        ((formatType == DatasetExporter.ExportFormat.Csv)
                             ? ","
                             : "	");
                }


            }


        }





        private static void XslValueOfForHeaders
            (XmlWriter writer, IList<string> sHeaders,
            ICollection<string> sFileds, DatasetExporter.ExportFormat formatType)
        {
// xsl:value-of for headers
            for (int i = 0; i < sHeaders.Count; i++)
            {
                writer.WriteString("\"");
                writer.WriteStartElement("xsl:value-of");
                writer.WriteAttributeString("select", "'" + sHeaders[i] + "'");
                writer.WriteEndElement(); // xsl:value-of
                writer.WriteString("\"");


                if (i != sFileds.Count - 1)
                    writer.WriteString
                        ((formatType == DatasetExporter.ExportFormat.Csv)
                             ? ","
                             : "	");
            }
        }

        private static void ConstructXslSkeleton(XmlTextWriter writer)
        {

            // xsl:stylesheet
            string ns = "http://www.w3.org/1999/XSL/Transform";
            writer.Formatting = Formatting.Indented;

            writer.WriteStartDocument();

            writer.WriteStartElement("xsl", "stylesheet", ns);
            writer.WriteAttributeString("version", "1.0");

            writer.WriteStartElement("xsl:output");
            writer.WriteAttributeString("method", "text");
            writer.WriteAttributeString("version", "4.0");

            writer.WriteEndElement();

        }
		



    }


}
