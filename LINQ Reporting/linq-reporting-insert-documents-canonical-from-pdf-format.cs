using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple class that holds a Document to be inserted via ReportingEngine.
    public class DocumentHolder
    {
        public Document Document { get; }

        public DocumentHolder(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source PDF file that will be inserted.
            string pdfPath = @"C:\Input\Source.pdf";

            // Load the PDF document. Aspose.Words can open PDF directly.
            Document pdfDocument = new Document(pdfPath);

            // Create a template Word document in memory.
            // The template contains a ReportingEngine tag that will be replaced with the PDF document.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // The tag syntax: <<doc [src.Document]>>
            builder.Writeln("<<doc [src.Document]>>");

            // Wrap the PDF document in a holder object so the ReportingEngine can access it via the "Document" property.
            DocumentHolder holder = new DocumentHolder(pdfDocument);

            // Build the report – this will replace the tag with the contents of the PDF document.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, new object[] { holder }, new[] { "src" });

            // Save the resulting document.
            string outputPath = @"C:\Output\Result.docx";
            template.Save(outputPath);
        }
    }
}
