using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingNumberedList
{
    // Root data model for the instruction manual.
    public class Manual
    {
        public List<Step> Steps { get; set; } = new();
    }

    // Individual step model.
    public class Step
    {
        public string Description { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data.
            string jsonPath = "steps.json";
            var sampleSteps = new Manual
            {
                Steps = new List<Step>
                {
                    new Step { Description = "Gather all required materials." },
                    new Step { Description = "Set up the workspace safely." },
                    new Step { Description = "Follow the assembly instructions step by step." },
                    new Step { Description = "Verify the final product for quality." }
                }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleSteps, Formatting.Indented));

            // Load JSON data into the model.
            Manual manual = JsonConvert.DeserializeObject<Manual>(File.ReadAllText(jsonPath))!;

            // Create the template document programmatically.
            string templatePath = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Instruction Manual");
            builder.Writeln();

            // Begin foreach loop over Steps.
            builder.Writeln("<<foreach [step in Steps]>>");
            // Numbered paragraph with restartNum to start numbering at 1 for each item.
            builder.Writeln("1. <<restartNum>><<[step.Description]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // Load the template for reporting.
            Document doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, manual, "manual");

            // Save the final document.
            string outputPath = "InstructionManual.docx";
            doc.Save(outputPath);
        }
    }
}
