using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample_page_1.jpg");

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document with long text to force line breaks.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to increase the chance of hyphenation.
        builder.PageSetup.PageWidth = 300; // points (~4.2 inches)
        builder.PageSetup.LeftMargin = 20;
        builder.PageSetup.RightMargin = 20;

        // Insert a long paragraph.
        builder.Font.Size = 12;
        builder.Writeln(
            "Antidisestablishmentarianism is a long word that often needs hyphenation. " +
            "Supercalifragilisticexpialidocious is another example of a word that can be split across lines. " +
            "The quick brown fox jumps over the lazy dog. " +
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
        );

        // Save the document as PDF (without hyphenation for now).
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF, enable automatic hyphenation.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Turn on automatic hyphenation.
        pdfDoc.HyphenationOptions.AutoHyphenation = true;
        // Optional: adjust hyphenation settings.
        pdfDoc.HyphenationOptions.HyphenationZone = 360; // default
        pdfDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        pdfDoc.HyphenationOptions.HyphenateCaps = true;

        // -----------------------------------------------------------------
        // Step 3: Render the first page of the PDF to an image.
        // -----------------------------------------------------------------
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            PageSet = new PageSet(0) // render only the first page (zero‑based index)
        };

        pdfDoc.Save(imagePath, imageOptions);

        // -----------------------------------------------------------------
        // Validation: ensure the image file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(imagePath))
            throw new InvalidOperationException($"Failed to create the image file at '{imagePath}'.");

        // The example finishes without requiring user interaction.
    }
}
