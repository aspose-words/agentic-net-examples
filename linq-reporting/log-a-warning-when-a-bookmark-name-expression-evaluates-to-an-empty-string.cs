using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingBookmarkWarning
{
    // Model class used by the LINQ Reporting template.
    public class ReportModel
    {
        // Bookmark name may be empty; initialize to empty string.
        public string BookmarkName { get; set; } = string.Empty;

        // Sample title to appear inside the bookmark.
        public string Title { get; set; } = "Default Title";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for possible data sources.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // Step 1: Create the template document with a bookmark tag.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a bookmark tag that uses the model's BookmarkName expression.
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");

            // Save the template to disk.
            doc.Save(templatePath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // Step 2: Load the template back for reporting.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Step 3: Prepare the data model.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                // Intentionally leave BookmarkName empty to trigger the warning.
                BookmarkName = string.Empty,
                Title = "Hello from LINQ Reporting"
            };

            // -----------------------------------------------------------------
            // Step 4: Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine
            {
                // Enable inline error messages so the engine does not throw on most errors.
                Options = ReportBuildOptions.InlineErrorMessages
            };

            bool success;
            try
            {
                // BuildReport returns true if the report was generated without errors.
                success = engine.BuildReport(loadedDoc, model, "model");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("bookmark's name"))
            {
                // The bookmark name evaluated to an empty string – log a warning and continue.
                Console.WriteLine("Warning: Bookmark name expression evaluated to an empty string.");
                success = false;
            }

            // If the engine reported failure (e.g., other errors), you could handle it here.
            if (!success && !engine.Options.HasFlag(ReportBuildOptions.InlineErrorMessages))
            {
                Console.WriteLine("Report generation completed with errors.");
            }

            // -----------------------------------------------------------------
            // Step 5: Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
