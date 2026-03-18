using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonReportExample
{
    static void Main()
    {
        // Path to the Word template that contains the reporting tags.
        // The template should reference the data source name "persons",
        // e.g. <<foreach [persons]>><<[Name]>><</foreach>>.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Load the template document (lifecycle rule: load).
        Document doc = new Document(templatePath);

        // JSON string that represents a root array.
        // This will be flattened by the reporting engine and treated as a table.
        string json = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Name"": ""Jane Smith"", ""Age"": 25 }
        ]";

        // Convert the JSON string to a memory stream.
        using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
        {
            // Create a JsonDataSource from the stream (lifecycle rule: create).
            JsonDataSource dataSource = new JsonDataSource(jsonStream);

            // Build the report using the data source.
            // The third argument is the name used in the template to reference the data.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "persons");
        }

        // Path where the generated report will be saved (lifecycle rule: save).
        string outputPath = @"C:\Reports\GeneratedReport.docx";

        // Save the populated document.
        doc.Save(outputPath);
    }
}
