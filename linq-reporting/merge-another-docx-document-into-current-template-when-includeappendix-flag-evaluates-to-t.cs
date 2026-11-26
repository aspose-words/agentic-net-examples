using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingMerge
{
    // Data model for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
        public bool IncludeAppendix { get; set; } = true; // Flag to control appendix merging.
        // Placeholder text that will be replaced by the appendix if needed.
        public string AppendixPlaceholder { get; set; } = "APPENDIX_CONTENT";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // File paths.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string appendixPath = Path.Combine(outputDir, "Appendix.docx");
            string resultPath = Path.Combine(outputDir, "Report.docx");

            // 1. Create the main template document with LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("=== Report ===");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("");
            // Conditional block that will be processed by the reporting engine.
            builder.Writeln("<<if [model.IncludeAppendix]>>");
            builder.Writeln("<<[model.AppendixPlaceholder]>>");
            builder.Writeln("<</if>>");
            templateDoc.Save(templatePath);

            // 2. Create a simple appendix document that will be merged when the flag is true.
            Document appendixDoc = new Document();
            DocumentBuilder appBuilder = new DocumentBuilder(appendixDoc);
            appBuilder.Writeln("=== Appendix ===");
            appBuilder.Writeln("This is the appended appendix content.");
            appendixDoc.Save(appendixPath);

            // 3. Prepare the data model.
            ReportModel model = new ReportModel
            {
                Title = "Quarterly Sales Report",
                IncludeAppendix = true // Change to false to skip merging.
            };

            // 4. Load the template and build the report using LINQ Reporting.
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. If the flag is true, merge the appendix document.
            if (model.IncludeAppendix)
            {
                Document appendixToMerge = new Document(appendixPath);
                // Append the appendix at the end of the report.
                reportDoc.AppendDocument(appendixToMerge, ImportFormatMode.KeepSourceFormatting);
            }

            // 6. Save the final report.
            reportDoc.Save(resultPath, SaveFormat.Docx);
        }
    }
}
