using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string dataFile = Path.Combine(workDir, "people.json");
        string templateFile = Path.Combine(workDir, "template.docx");
        string resultFile = Path.Combine(workDir, "report.docx");

        // 1. Create sample JSON data (array of objects).
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30 },
  { ""Name"": ""Bob"",   ""Age"": 25 },
  { ""Name"": ""Carol"", ""Age"": 28 }
]";
        File.WriteAllText(dataFile, jsonContent, Encoding.UTF8);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [in jsonData]>>");
        builder.Writeln("- <<[Name]>> (Age: <<[Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFile);

        // 3. Load the template back (as required by the lifecycle rule).
        Document loadedTemplate = new Document(templateFile);

        // 4. Create a JsonDataSource from the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(dataFile);

        // 5. Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(loadedTemplate, jsonDataSource, "jsonData");

        // 6. Save the generated report.
        loadedTemplate.Save(resultFile);
    }
}
