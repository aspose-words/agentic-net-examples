using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // When null the signature block will be omitted.
        public string? Signatory { get; set; } = null;

        // Additional sample property.
        public string Title { get; set; } = "Sample Report";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a blank document and a builder to compose the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Simple title.
            builder.Writeln($"<<[model.Title]>>");
            builder.Writeln();

            // Conditional block: include the signature line only when Signatory is not null.
            builder.Writeln("<<if [model.Signatory != null]>>");
            builder.Writeln("Signature:");
            // Insert a signature line with static options.
            builder.InsertSignatureLine(new SignatureLineOptions
            {
                Signer = "Signer",
                SignerTitle = "Title",
                ShowDate = true,
                DefaultInstructions = false,
                Instructions = "Please sign here."
            });
            // Show the name of the signatory.
            builder.Writeln("Signed by: <<[model.Signatory]>>");
            builder.Writeln("<</if>>");

            // Build the report.
            ReportModel model = new ReportModel
            {
                Title = "Quarterly Financial Summary",
                Signatory = "John Doe" // Set to null to see the block omitted.
            };

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            string outPath = Path.Combine(outputDir, "ConditionalSignatureReport.docx");
            doc.Save(outPath);
        }
    }
}
