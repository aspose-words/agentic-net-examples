using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    // Simple data model for JSON serialization.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using (MemoryStream templateStream = new MemoryStream())
        {
            templateDoc.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0; // Reset for reading.

            // 2. Load the template from the stream.
            Document reportDoc = new Document(templateStream);

            // 3. Prepare JSON data and create a JsonDataSource from a stream.
            List<Person> persons = new()
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 }
            };
            string json = JsonConvert.SerializeObject(persons);
            using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                jsonStream.Position = 0; // Ensure the stream is at the beginning.
                JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

                // 4. Build the report using the ReportingEngine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(reportDoc, jsonDataSource, "persons");

                // 5. Save the generated report as RTF.
                string outputPath = "Report.rtf";
                reportDoc.Save(outputPath, SaveFormat.Rtf);
            }
        }
    }
}
