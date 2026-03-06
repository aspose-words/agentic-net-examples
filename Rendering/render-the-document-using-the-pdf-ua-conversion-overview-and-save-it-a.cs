using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace RenderPdfUaExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the folder that contains the source document and where the PDF will be saved.
            string dataDir = @"C:\Data\";

            // Load the source Word document.
            Document doc = new Document(dataDir + "input.docx");

            // Create PDF save options and configure them for PDF/UA compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the compliance level to PDF/UA‑1 (ISO 14289‑1).
                Compliance = PdfCompliance.PdfUa1,

                // PDF/UA requires the document title to be shown in the viewer’s title bar.
                DisplayDocTitle = true,

                // Export the document structure (tags) which is mandatory for PDF/UA.
                ExportDocumentStructure = true
            };

            // Save the document as a PDF file with the specified options.
            doc.Save(dataDir + "output.pdf", pdfOptions);

            Console.WriteLine("PDF/UA document saved successfully to " + dataDir + "output.pdf");
        }
    }
}
