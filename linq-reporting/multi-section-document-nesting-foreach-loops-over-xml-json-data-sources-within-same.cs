using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the Word template that contains nested <<foreach>> tags for XML and JSON.
        Document template = new Document("Template.docx");

        // Create the XML data source (e.g., a file with a list of people).
        XmlDataSource xmlSource = new XmlDataSource("People.xml");

        // Create the JSON data source (e.g., a file with a list of orders).
        JsonDataSource jsonSource = new JsonDataSource("Orders.json");

        // Build the report using both data sources.
        // The first source can be referenced in the template as "xml",
        // the second as "json" (e.g., <<foreach [xml/Person]>> and <<foreach [json/Order]>>).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(
            template,
            new object[] { xmlSource, jsonSource },
            new string[] { "xml", "json" });

        // Save the populated document.
        template.Save("Result.docx");
    }
}
