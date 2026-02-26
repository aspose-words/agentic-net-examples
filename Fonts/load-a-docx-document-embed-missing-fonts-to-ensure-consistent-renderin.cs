using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCX file path
        string inputPath = @"C:\Docs\input.docx";

        // Output PDF file path
        string outputPath = @"C:\Docs\output.pdf";

        // Load the DOCX document
        Document doc = new Document(inputPath);

        // Configure PDF save options to embed all fonts used in the document
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed every font (including standard Windows fonts) into the PDF
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            // Embed the full font files (no subsetting) to guarantee all glyphs are present
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified embedding options
        doc.Save(outputPath, pdfOptions);
    }
}
