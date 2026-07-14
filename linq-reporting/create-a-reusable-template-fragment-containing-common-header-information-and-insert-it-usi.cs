using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingIncludeExample
{
    // Data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; } = "";
        public string Body { get; set; } = "";
    }

    // Wrapper for the header document so it can be referenced by the <<doc>> tag.
    public class HeaderWrapper
    {
        public Document Document { get; set; } = null!;
    }

    public class Program
    {
        public static void Main()
        {
            // Working directory – the example runs in a writable folder.
            string workDir = Directory.GetCurrentDirectory();

            // -----------------------------------------------------------------
            // 1. Create a reusable header fragment (HeaderTemplate.docx).
            // -----------------------------------------------------------------
            string headerPath = Path.Combine(workDir, "HeaderTemplate.docx");
            Document headerDoc = new Document();
            DocumentBuilder headerBuilder = new DocumentBuilder(headerDoc);

            // Header content with LINQ Reporting tags.
            headerBuilder.Writeln("<<[model.Title]>>");
            headerBuilder.Writeln("------------------------------");
            headerDoc.Save(headerPath);

            // Load the header document so it can be passed to the <<doc>> tag.
            HeaderWrapper headerWrapper = new HeaderWrapper
            {
                Document = new Document(headerPath)
            };

            // -----------------------------------------------------------------
            // 2. Create the main template that includes the header fragment.
            // -----------------------------------------------------------------
            string mainPath = Path.Combine(workDir, "MainTemplate.docx");
            Document mainDoc = new Document();
            DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

            // Include the previously created header fragment using the supported <<doc>> tag.
            // The tag expects a data source named "src" with a public Document property.
            mainBuilder.Writeln("<<doc [src.Document]>>");
            // Body content with a tag.
            mainBuilder.Writeln("<<[model.Body]>>");
            mainDoc.Save(mainPath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Title = "Monthly Sales Report",
                Body = "This month we achieved a 15% increase in revenue."
            };

            // -----------------------------------------------------------------
            // 4. Load the main template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(mainPath);
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using two data sources: the model (named "model")
            // and the header wrapper (named "src").
            engine.BuildReport(reportDoc,
                new object[] { model, headerWrapper },
                new[] { "model", "src" });

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(workDir, "ReportOutput.docx");
            reportDoc.Save(outputPath);
        }
    }
}
