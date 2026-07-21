using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Register code page provider for .NET Core compatibility.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Define file paths.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinqReportingDemo");
            Directory.CreateDirectory(workDir);
            string jsonPath = Path.Combine(workDir, "data.json");
            string templatePath = Path.Combine(workDir, "template.docx");
            string reportPath = Path.Combine(workDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create sample JSON data with two numeric properties.
            // -----------------------------------------------------------------
            string jsonContent = @"
[
    { ""Value1"": 10, ""Value2"": 5 },
    { ""Value1"": 20, ""Value2"": 7 },
    { ""Value1"": 15, ""Value2"": 3 }
]";
            File.WriteAllText(jsonPath, jsonContent.Trim());

            // -----------------------------------------------------------------
            // 2. Build a Word template that uses LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("LINQ Reporting Demo – Calculated Field");
            builder.Writeln();

            // Begin foreach loop over the JSON array (named 'items' in the BuildReport call).
            builder.Writeln("<<foreach [item in items]>>");
            // Output each item's values and the calculated sum.
            builder.Writeln("Value1: <<[item.Value1]>>, Value2: <<[item.Value2]>>, Sum: <<[item.Value1 + item.Value2]>>");
            // End foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the JSON data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root name 'items' matches the foreach variable.
            engine.BuildReport(reportDoc, jsonDataSource, "items");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Optional: indicate completion (no interactive input).
            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}
