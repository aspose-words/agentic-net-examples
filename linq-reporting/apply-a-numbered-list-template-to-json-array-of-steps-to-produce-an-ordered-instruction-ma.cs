using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Root data model matching the JSON structure.
    public class Manual
    {
        public List<Step> Steps { get; set; } = new();
    }

    public class Step
    {
        public string Description { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data.
            string jsonContent = @"{
  ""Steps"": [
    { ""Description"": ""Gather all necessary materials."" },
    { ""Description"": ""Prepare the workspace and tools."" },
    { ""Description"": ""Assemble the components in order."" },
    { ""Description"": ""Test the assembled product."" },
    { ""Description"": ""Clean up the area."" }
  ]
}";
            // Write JSON to a local file.
            const string jsonPath = "steps.json";
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Deserialize JSON into the data model.
            Manual manual = JsonConvert.DeserializeObject<Manual>(jsonContent)!;

            // Create the LINQ Reporting template programmatically.
            const string templatePath = "template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Title.
            builder.Writeln("Instruction Manual");
            builder.Writeln(); // Blank line.

            // Begin foreach block over the Steps collection.
            builder.Writeln("<<foreach [step in Steps]>>");

            // Apply numbering to each step description.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("<<[step.Description]>>");
            builder.ListFormat.RemoveNumbers();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new(templatePath);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, manual, "manual");

            // Save the final document.
            const string outputPath = "InstructionManual.docx";
            reportDoc.Save(outputPath);
        }
    }
}
