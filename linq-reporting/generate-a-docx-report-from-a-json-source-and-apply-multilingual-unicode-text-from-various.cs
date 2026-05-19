using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Enable full Unicode support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string jsonPath = Path.Combine(outputDir, "data.json");
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Sample JSON containing multilingual records.
        string jsonContent = @"[
  {
    ""Name"": ""John Doe"",
    ""Greeting"": ""Hello"",
    ""Language"": ""English"",
    ""Message"": ""Welcome""
  },
  {
    ""Name"": ""Иван Иванов"",
    ""Greeting"": ""Привет"",
    ""Language"": ""Russian"",
    ""Message"": ""Добро пожаловать""
  },
  {
    ""Name"": ""张伟"",
    ""Greeting"": ""你好"",
    ""Language"": ""Chinese"",
    ""Message"": ""欢迎""
  },
  {
    ""Name"": ""محمد علي"",
    ""Greeting"": ""مرحبا"",
    ""Language"": ""Arabic"",
    ""Message"": ""أهلا وسهلا""
  },
  {
    ""Name"": ""राहुल शर्मा"",
    ""Greeting"": ""नमस्ते"",
    ""Language"": ""Hindi"",
    ""Message"": ""स्वागत है""
  }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // -----------------------------
        // Create the template document.
        // -----------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Multilingual Report");
        builder.Writeln();

        // Begin foreach loop over the JSON collection named 'persons'.
        builder.Writeln("<<foreach [p in persons]>>");

        // Build a table inside the loop.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Name");
        builder.InsertCell(); builder.Writeln("Greeting");
        builder.InsertCell(); builder.Writeln("Language");
        builder.InsertCell(); builder.Writeln("Message");
        builder.EndRow();

        // Data row with LINQ Reporting tags.
        builder.InsertCell(); builder.Writeln("<<[p.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Greeting]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Language]>>");
        builder.InsertCell(); builder.Writeln("<<[p.Message]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------
        // Generate the final report.
        // -----------------------------
        Document reportDoc = new Document(templatePath);

        // Load JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Configure and run the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the populated report.
        reportDoc.Save(reportPath);
    }
}
