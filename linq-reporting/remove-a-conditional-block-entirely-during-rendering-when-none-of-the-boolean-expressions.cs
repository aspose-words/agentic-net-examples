using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalRemoval
{
    // Data model used by the template.
    public class ReportModel
    {
        public bool ShowA { get; set; } = false;
        public bool ShowB { get; set; } = false;
        public string Message { get; set; } = "No sections displayed.";
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with two conditional blocks.
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Conditional block A
            builder.Writeln("<<if [model.ShowA]>>");
            builder.Writeln("Section A is visible.");
            builder.Writeln("<</if>>");

            // Conditional block B
            builder.Writeln("<<if [model.ShowB]>>");
            builder.Writeln("Section B is visible.");
            builder.Writeln("<</if>>");

            // Fallback message when no blocks are rendered.
            builder.Writeln("<<[model.Message]>>");

            templateDoc.Save(templatePath);

            // 2. Load the template for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare the data source where both conditions are false.
            var model = new ReportModel
            {
                ShowA = false,
                ShowB = false,
                Message = "All conditional sections were removed."
            };

            // 4. Build the report with the RemoveEmptyParagraphs option.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(doc, model, "model");

            // 5. Save the final document.
            var outputPath = "Output.docx";
            doc.Save(outputPath);

            // Optional: display the resulting text in the console.
            Console.WriteLine("Report generated. Content:");
            Console.WriteLine(doc.GetText());
        }
    }
}
