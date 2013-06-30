using System;
using System.Collections.Generic;
using System.Text;
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
            string[] sFileds, DatasetExporter.ExportFormat FormatType)
        {


            try
            {

                ConstructXslSkeleton(writer);

                // xsl-template
                writer.WriteStartElement("xsl:template");
                writer.WriteAttributeString("match", "/");


                // xsl:value-of for headers
                for (int i = 0; i < sHeaders.Length; i++)
                {
                    writer.WriteString("\"");
                    writer.WriteStartElement("xsl:value-of");
                    writer.WriteAttributeString("select", "'" + sHeaders[i] + "'");
                    writer.WriteEndElement(); // xsl:value-of
                    writer.WriteString("\"");


                    if (i != sFileds.Length - 1)
                        writer.WriteString
                            ((FormatType == DatasetExporter.ExportFormat.CSV)
                            ? "," : "	");


                }



                // xsl:for-each
                writer.WriteStartElement("xsl:for-each");
                writer.WriteAttributeString("select", "Export/Values");
                writer.WriteString("\r\n");



                // xsl:value-of for data fields
                for (int i = 0; i < sFileds.Length; i++)
                {


                    writer.WriteString("\"");
                    writer.WriteStartElement("xsl:value-of");
                    writer.WriteAttributeString("select", sFileds[i]);
                    writer.WriteEndElement(); // xsl:value-of
                    writer.WriteString("\"");

                    if (i != sFileds.Length - 1)
                    {

                        writer.WriteString
                            ((FormatType == DatasetExporter.ExportFormat.CSV)
                            ? "," : "	");

                    }

                }


                writer.WriteEndElement(); // xsl:for-each
                writer.WriteEndElement(); // xsl-template
                writer.WriteEndElement(); // xsl:stylesheet
                writer.WriteEndDocument();


            }
            catch (Exception Ex)
            {
                throw Ex;
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
