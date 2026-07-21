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
        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple report header.
        builder.Writeln("People Report");
        builder.Writeln();

        // LINQ Reporting tags: iterate over the JSON array named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using (MemoryStream templateStream = new MemoryStream())
        {
            templateDoc.Save(templateStream, SaveFormat.Docx);
            templateStream.Position = 0; // Reset for reading.

            // ---------- Load the template from the stream ----------
            Document reportDoc = new Document(templateStream);

            // ---------- Prepare JSON data ----------
            string json = @"[
                { ""Name"": ""John Doe"", ""Age"": 30 },
                { ""Name"": ""Jane Smith"", ""Age"": 25 },
                { ""Name"": ""Bob Johnson"", ""Age"": 40 }
            ]";

            // Create a JsonDataSource from the JSON string via a memory stream.
            using (MemoryStream jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                jsonStream.Position = 0; // Ensure the stream is at the beginning.

                JsonDataSource jsonDataSource = new JsonDataSource(jsonStream);

                // ---------- Build the report ----------
                ReportingEngine engine = new ReportingEngine();
                // The root name "persons" matches the tag used in the template.
                engine.BuildReport(reportDoc, jsonDataSource, "persons");

                // ---------- Save the generated report as RTF ----------
                reportDoc.Save("PeopleReport.rtf", SaveFormat.Rtf);
            }
        }
    }
}
