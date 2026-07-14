using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;          // JsonDataSource resides in this namespace

public class Program
{
    public static void Main()
    {
        // 1. Prepare JSON data.
        string json = @"[
            { ""Name"": ""John Doe"", ""Age"": 30 },
            { ""Name"": ""Jane Smith"", ""Age"": 25 }
        ]";

        // Write JSON to a memory stream.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(json));

        // 2. Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using var templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0; // Reset for reading.

        // 3. Load the template document from the stream.
        var doc = new Document(templateStream);

        // 4. Create a JsonDataSource from the JSON stream.
        jsonStream.Position = 0; // Ensure the stream is at the beginning.
        var jsonDataSource = new JsonDataSource(jsonStream);

        // 5. Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "persons");

        // 6. Save the generated report as RTF.
        doc.Save("Report.rtf", SaveFormat.Rtf);
    }
}
