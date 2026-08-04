using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public partial class Program
{
    public static void Main()
    {
        // Register code page provider for possible encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths for the sample JSON, template, and output report.
        const string jsonPath = "data.json";
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // 1. Create sample JSON data (array of person objects) and write it to a file.
        string jsonContent = @"[
            { ""Name"": ""John Doe"", ""Age"": 30, ""Email"": ""john@example.com"" },
            { ""Name"": ""Jane Smith"", ""Age"": 25, ""Email"": ""jane@example.com"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // LINQ Reporting foreach tag iterating over the JSON root collection named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name : <<[person.Name]>>");
        builder.Writeln("Age  : <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template back (required by the lifecycle rule).
        Document doc = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name is "persons" – it must match the name used in the template tags.
        engine.BuildReport(doc, jsonData, "persons");

        // 6. Save the generated report.
        doc.Save(outputPath);
    }
}
