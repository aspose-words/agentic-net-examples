using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonReportExample
{
    static void Main()
    {
        // JSON string that will be used as the data source.
        string json = @"{
            ""persons"": [
                { ""Name"": ""John"", ""Age"": 30 },
                { ""Name"": ""Jane"", ""Age"": 25 }
            ]
        }";

        // Convert the JSON string to a memory stream.
        using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
        {
            // Create a JsonDataSource from the JSON stream.
            JsonDataSource dataSource = new JsonDataSource(jsonStream);

            // Create a blank document that will serve as the template.
            Document template = new Document();

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("persons") matches the top‑level JSON array.
            engine.BuildReport(template, dataSource, "persons");

            // Save the resulting document to a memory stream.
            using (MemoryStream output = new MemoryStream())
            {
                template.Save(output, SaveFormat.Docx);

                // The output stream now contains the generated DOCX file.
                // For demonstration, we can write its length to the console.
                Console.WriteLine($"Report generated, size = {output.Length} bytes");
            }
        }
    }
}
