using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add a paragraph with a run of text.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);
        Run run = new Run(doc, "Hello, world with embedded font!");
        paragraph.AppendChild(run);

        // Set the font for the run using Aspose.Words.Font.
        run.Font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (run.Font.Name != "Arial")
        {
            throw new InvalidOperationException("Font name was not set correctly.");
        }

        // Configure PDF save options to embed all used fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        string outputPath = "EmbeddedFontDocument.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Failed to create the output PDF.", outputPath);
        }

        // Indicate successful completion.
        Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
    }
}
