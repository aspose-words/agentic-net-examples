using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with text that typically uses OpenType ligatures (e.g., "ff", "fi").
        Paragraph paragraph = new Paragraph(doc);
        Run run = new Run(doc, "Office");
        // Use a font that contains OpenType features.
        run.Font.Name = "Times New Roman";
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Render the document to PDF.
        string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("PDF file was not created.");
        }

        // Simple inspection of the PDF content for embedded font markers.
        byte[] pdfBytes = File.ReadAllBytes(pdfPath);
        string pdfText = Encoding.ASCII.GetString(pdfBytes);
        bool hasEmbeddedFont = pdfText.Contains("/FontFile") ||
                               pdfText.Contains("/FontFile2") ||
                               pdfText.Contains("/FontFile3");

        Console.WriteLine(hasEmbeddedFont
            ? "PDF generated successfully with embedded fonts."
            : "PDF generated, but no embedded font markers were found.");
    }
}
