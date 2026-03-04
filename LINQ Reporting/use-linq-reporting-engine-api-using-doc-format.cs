using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains LINQ Reporting Engine tags.
        Document template = new Document("Template.docx");

        // Create a JSON data source. The file "data.json" should contain the data
        // referenced in the template (e.g. <<[persons.Name]>>).
        JsonDataSource dataSource = new JsonDataSource("data.json");

        // Populate the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "persons");

        // Save the generated report as a DOC file.
        template.Save("Report.docx");
    }
}
