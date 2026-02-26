using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the PDF template that contains LINQ Reporting tags, e.g. <<[Data.Name]>>.
            Document template = new Document(@"C:\Templates\ReportTemplate.pdf");

            // Create a simple anonymous data source that matches the tags in the template.
            var dataSource = new
            {
                Name = "John Doe",
                Age = 30,
                Address = "123 Main St, Anytown"
            };

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Example option: remove empty paragraphs after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the overload that allows referencing the data source object itself.
            // The third argument ("Data") is the name used inside the template tags.
            engine.BuildReport(template, dataSource, "Data");

            // Dynamically insert an additional document (e.g., a terms‑and‑conditions section).
            Document extraDoc = new Document(@"C:\Templates\TermsAndConditions.docx");
            DocumentBuilder builder = new DocumentBuilder(template);
            // Move the cursor to the end of the main story and insert the extra document.
            builder.MoveToDocumentEnd();
            builder.InsertDocument(extraDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the final report as PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: set high‑quality rendering for better appearance.
                UseHighQualityRendering = true
            };
            template.Save(@"C:\Output\FinalReport.pdf", pdfOptions);
        }
    }
}
