using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Folder that contains the source PDF files.
        string pdfFolder = @"C:\InputPdfs";

        // Load each PDF as an Aspose.Words Document and wrap it in a simple class.
        // The wrapper is needed so the ReportingEngine can reference the Document property.
        PdfWrapper[] pdfSources = Directory.GetFiles(pdfFolder, "*.pdf")
            .Select(path => new Document(path))               // Load PDF into a Document.
            .Select(doc => new PdfWrapper { Document = doc }) // Wrap the Document.
            .ToArray();

        // Create a template document that defines a repeatable region.
        // The <<foreach [src]>> tag iterates over the array passed to BuildReport.
        // Inside the region we insert each PDF using the <<doc [src.Document]>> tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Combined PDFs:");
        builder.Writeln("<<foreach [src]>>");
        builder.Writeln("<<doc [src.Document]>>");
        builder.Writeln("<</foreach>>");

        // Build the report by providing the array of wrappers and the data source name "src".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, pdfSources, new[] { "src" });

        // Save the merged document.
        template.Save(@"C:\Output\Combined.docx");
    }

    // Simple wrapper class exposing a Document property for the ReportingEngine.
    public class PdfWrapper
    {
        public Document Document { get; set; }
    }
}
