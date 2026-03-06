using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace MergeTemplateExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that will serve as the base for the merge template.
            const string sourceDocPath = @"C:\Docs\SourceTemplate.docx";

            // Load the existing DOCX document using the Document(string) constructor (load rule).
            Document templateDocument = new Document(sourceDocPath);

            // -----------------------------------------------------------------
            // Create a merge template model.
            // Here we add a simple MERGEFIELD that can be populated later.
            // This uses DocumentBuilder, which is a standard way to modify a document.
            // -----------------------------------------------------------------
            DocumentBuilder builder = new DocumentBuilder(templateDocument);
            // Move the cursor to the end of the document (or any desired location).
            builder.MoveToDocumentEnd();
            // Insert a merge field named "CustomerName".
            builder.InsertField("MERGEFIELD CustomerName ");

            // Optionally, you could prepare the template for reporting engine usage.
            // For example, you could populate static parts of the template now.
            // ReportingEngine can be used later to generate concrete documents from this template.
            // ReportingEngine engine = new ReportingEngine();
            // engine.BuildReport(templateDocument, new { CustomerName = "<<[CustomerName]>>" });

            // -----------------------------------------------------------------
            // Save the prepared template as a PDF file.
            // The Save(string, SaveFormat) overload automatically determines the format
            // from the SaveFormat enum (save rule).
            // -----------------------------------------------------------------
            const string outputPdfPath = @"C:\Docs\Result.pdf";
            templateDocument.Save(outputPdfPath, SaveFormat.Pdf);
        }
    }
}
