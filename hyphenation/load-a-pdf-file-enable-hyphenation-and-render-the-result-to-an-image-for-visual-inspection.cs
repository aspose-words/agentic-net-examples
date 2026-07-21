using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a sample document with long words that can be hyphenated.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document as a PDF – this will be the source PDF we later load.
        const string pdfPath = "sample.pdf";
        doc.Save(pdfPath);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF.");

        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate words in the document.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary registration failed.");

        // Load the previously saved PDF.
        var pdfDoc = new Document(pdfPath);

        // Enable automatic hyphenation for the loaded document.
        pdfDoc.HyphenationOptions.AutoHyphenation = true;
        // Use a valid hyphenation zone (default is 360 = 0.25 inch).
        pdfDoc.HyphenationOptions.HyphenationZone = 360;
        pdfDoc.HyphenationOptions.HyphenateCaps = true;
        pdfDoc.HyphenationOptions.ConsecutiveHyphenLimit = 0; // 0 = unlimited

        // Render the first page of the PDF to an image for visual inspection.
        var imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            PageSet = new PageSet(0), // Render only the first page (zero‑based index).
            Resolution = 300
        };
        const string imagePath = "rendered.jpg";
        pdfDoc.Save(imagePath, imageOptions);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Failed to render the PDF to an image.");
    }
}
