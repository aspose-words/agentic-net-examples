using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output paths
        string docPath = "SampleDocument.docx";
        string pdfPath = "SampleDocument.pdf";

        // Create a new document
        Document doc = new Document();

        // Add a paragraph with some text
        Paragraph para = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para);
        Run run = new Run(doc, "This text uses the Arial font.");
        para.AppendChild(run);

        // Set the font for the run
        Aspose.Words.Font font = run.Font;
        font.Name = "Arial";

        // Validate that the font name was set correctly
        if (font.Name != "Arial")
        {
            throw new InvalidOperationException("Failed to set the font name.");
        }

        // Save the document as DOCX (optional, just to demonstrate saving)
        doc.Save(docPath, SaveFormat.Docx);

        // Prepare PDF save options with font embedding enabled
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed all fonts used in the document
            EmbedFullFonts = true,
            // Ensure that fonts are embedded even if they are standard
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF file was not created.", pdfPath);
        }

        // Clean up temporary DOCX file (optional)
        if (File.Exists(docPath))
        {
            File.Delete(docPath);
        }
    }
}
