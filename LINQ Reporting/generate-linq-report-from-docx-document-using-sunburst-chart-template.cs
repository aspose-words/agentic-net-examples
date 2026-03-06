using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace SunburstReportExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains a Sunburst chart with LINQ placeholders.
            const string templatePath = "TemplateSunburst.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare hierarchical XML data that matches the placeholders used in the template.
            // The XML structure should reflect the hierarchy required by the Sunburst chart
            // (e.g., Region → Country → Department → Value).
            string xmlData = @"
                <Sales>
                    <Item>
                        <Region>Europe</Region>
                        <Country>UK</Country>
                        <Department>London Dep.</Department>
                        <Value>1236</Value>
                    </Item>
                    <Item>
                        <Region>Europe</Region>
                        <Country>UK</Country>
                        <Department>Liverpool Dep.</Department>
                        <Value>851</Value>
                    </Item>
                    <Item>
                        <Region>Europe</Region>
                        <Country>France</Country>
                        <Department>Paris Dep.</Department>
                        <Value>468</Value>
                    </Item>
                    <Item>
                        <Region>North America</Region>
                        <Country>USA</Country>
                        <Department>Denver Dep.</Department>
                        <Value>527</Value>
                    </Item>
                    <Item>
                        <Region>North America</Region>
                        <Country>Canada</Country>
                        <Department>Toronto Dep.</Department>
                        <Value>457</Value>
                    </Item>
                    <Item>
                        <Region>Oceania</Region>
                        <Country>Australia</Country>
                        <Department>Sydney Dep.</Department>
                        <Value>761</Value>
                    </Item>
                </Sales>";

            // Convert the XML string to a stream and create an XmlDataSource.
            using (MemoryStream xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlData)))
            {
                XmlDataSource dataSource = new XmlDataSource(xmlStream);

                // Create the reporting engine.
                ReportingEngine engine = new ReportingEngine();

                // Build the report by populating the template with the XML data.
                // The third argument ("sales") is the name used in the template to reference the data source.
                engine.BuildReport(doc, dataSource, "sales");
            }

            // Save the populated document with the Sunburst chart rendered.
            const string outputPath = "SunburstReport.docx";
            doc.Save(outputPath);
        }
    }
}
