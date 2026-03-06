using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace RepeatingSectionExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains a repeating section (SDT type RepeatingSection).
            Document template = new Document("Template.docx");

            // XML data that will be mapped to the repeating section.
            // The XPath used in the template should match the structure below.
            string xmlData = @"
                <books>
                    <book>
                        <title>Everyday Italian</title>
                        <author>Giada De Laurentiis</author>
                    </book>
                    <book>
                        <title>The C Programming Language</title>
                        <author>Brian W. Kernighan, Dennis M. Ritchie</author>
                    </book>
                    <book>
                        <title>Learning XML</title>
                        <author>Erik T. Ray</author>
                    </book>
                </books>";

            // Create an XmlDataSource from the XML string.
            XmlDataSource dataSource = new XmlDataSource(xmlData);

            // Build the report – this populates the repeating section with the XML data.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource);

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
