using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string jsonPath = Path.Combine(workDir, "data.json");
        string outputPath = Path.Combine(workDir, "report.docx");

        // 1. Create a simple JSON file with ISO‑8601 date.
        string jsonContent = @"{
    ""OrderDate"": ""2023-08-15T14:30:00Z"",
    ""CustomerName"": ""John Doe""
}";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build a Word template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order date (ISO‑8601): <<[order.OrderDate]>>");
        templateDoc.Save(templatePath);

        // 3. Load the template back.
        Document doc = new Document(templatePath);

        // 4. Configure JSON loading to recognise ISO‑8601 dates.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            // Explicit ISO‑8601 format (can be omitted because it is default, but shown for clarity).
            ExactDateTimeParseFormats = new List<string> { "yyyy-MM-ddTHH:mm:ssZ" },
            AlwaysGenerateRootObject = true
        };

        // 5. Create a JsonDataSource from the file.
        JsonDataSource jsonData = new JsonDataSource(jsonPath, loadOptions);

        // 6. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this scenario.
        engine.BuildReport(doc, jsonData, "order");

        // 7. Save the generated report.
        doc.Save(outputPath);
    }
}
