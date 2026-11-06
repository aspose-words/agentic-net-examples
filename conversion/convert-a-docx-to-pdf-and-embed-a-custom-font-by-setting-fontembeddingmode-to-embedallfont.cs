using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document using a custom font.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // The enum value for embedding all fonts may be named EmbedAll in some library versions.
        // Adjust the value according to the available enum members.
        saveOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

        // Save as PDF with embedded fonts.
        doc.Save("output.pdf", saveOptions);

        // Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
