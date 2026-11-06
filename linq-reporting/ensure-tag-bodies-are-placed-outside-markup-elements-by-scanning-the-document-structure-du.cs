using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

namespace LinqReportingTagPreprocess
{
    // Simple data model for the report.
    public class ReportModel
    {
        public string BookmarkName { get; set; } = "MyBookmark";
        public string Title { get; set; } = "Report Title";
    }

    // Visitor that scans the document for LINQ Reporting tags.
    public class TagPreprocessor : DocumentVisitor
    {
        public override VisitorAction VisitRun(Run run)
        {
            string text = run.Text;
            if (text.Contains("<<bookmark") && text.Contains("<</bookmark>>"))
            {
                // Example check: ensure the tag body (the text between opening and closing tags)
                // is not mixed with other markup in the same run.
                // For demonstration we simply output a diagnostic message.
                Console.WriteLine("Found bookmark tag in run: \"{0}\"", text);
            }
            return VisitorAction.Continue;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Step 1: Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a bookmark tag with its body placed outside any markup.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template for preprocessing.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Scan the document structure to ensure tag bodies are correctly placed.
            doc.Accept(new TagPreprocessor());

            // -------------------------------------------------
            // Step 3: Prepare the data model.
            // -------------------------------------------------
            ReportModel model = new ReportModel();

            // -------------------------------------------------
            // Step 4: Build the report using LINQ Reporting Engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // Step 5: Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
