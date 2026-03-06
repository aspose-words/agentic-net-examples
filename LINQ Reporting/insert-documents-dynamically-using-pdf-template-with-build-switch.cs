using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the Word template that contains the build switch, e.g. <<doc [src.Document] -sourceNumbering>>
        string templatePath = "Template.docx";

        // Path to the document that will be inserted dynamically
        string sourceDocPath = "Insert.docx";

        // Path where the final PDF will be saved
        string outputPdfPath = "Result.pdf";

        // Load the template document
        Document template = new Document(templatePath);

        // Load the source document to be inserted
        Document sourceDoc = new Document(sourceDocPath);

        // Create a ReportingEngine instance
        ReportingEngine engine = new ReportingEngine();

        // Build the report: the source document is referenced in the template by the name "src"
        // The template uses the "-sourceNumbering" switch to keep the inserted document's numbering as is
        engine.BuildReport(template, new object[] { sourceDoc }, new[] { "src" });

        // Save the populated document as PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions(); // default options
        template.Save(outputPdfPath, pdfOptions);
    }
}
