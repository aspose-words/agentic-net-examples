using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfSubsetExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string docxPath = @"C:\Input\SampleDocument.docx";

            // Path where the resulting PDF will be saved.
            string pdfPath = @"C:\Output\SampleDocument_Subset.pdf";

            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure that only the glyphs used in the document are embedded.
            // The default value of EmbedFullFonts is false, which enables subsetting.
            // Explicitly set it for clarity.
            pdfOptions.EmbedFullFonts = false;

            // Embed all fonts that are used in the document (default behavior).
            pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

            // Save the document as PDF using the configured options.
            doc.Save(pdfPath, pdfOptions);
        }
    }
}
