using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple report header
        builder.Writeln("Report of Persons:");
        // LINQ Reporting foreach tag iterating over the JSON array named "persons"
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream
        using (MemoryStream templateStream = new MemoryStream())
        {
            templateDoc.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0; // Reset for reading

            // ---------- Load the template from the stream ----------
            Document reportDoc = new Document(templateStream);

            // ---------- Prepare JSON data source ----------
            string json = @"[
                { ""Name"": ""Alice"", ""Age"": 30 },
                { ""Name"": ""Bob"",   ""Age"": 25 },
                { ""Name"": ""Carol"", ""Age"": 28 }
            ]";

            byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
            using (MemoryStream jsonStream = new MemoryStream(jsonBytes))
            {
                JsonDataSource dataSource = new JsonDataSource(jsonStream);

                // ---------- Build the report ----------
                ReportingEngine engine = new ReportingEngine();
                // The root name "persons" matches the tag used in the template
                engine.BuildReport(reportDoc, dataSource, "persons");

                // ---------- Save the final report as RTF ----------
                reportDoc.Save("Report.rtf", SaveFormat.Rtf);
            }
        }
    }
}
