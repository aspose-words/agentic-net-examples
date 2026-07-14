using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class HyphenationDemo
{
    public static void Main()
    {
        // Paths for temporary files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string pdfPath = "sample.pdf";
        const string imagePath = "rendered.jpg";

        // 1. Create a minimal hyphenation dictionary for English (US).
        // The dictionary format is OpenOffice style: first line is encoding, then word=hyphenation points.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // 2. Build a source document with long words that require hyphenation.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Use a narrow page width to force line wrapping.
        sourceDoc.FirstSection.PageSetup.PageWidth = 200; // points
        sourceDoc.FirstSection.PageSetup.LeftMargin = 20;
        sourceDoc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing words that match the dictionary.
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // 3. Save the source document as PDF – this will be the file we later load.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");

        // 4. Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // 5. Enable automatic hyphenation for the loaded document.
        pdfDoc.HyphenationOptions.AutoHyphenation = true;
        // Use the default hyphenation zone (360 = 0.25 inch) – 0 is invalid.
        pdfDoc.HyphenationOptions.HyphenationZone = 360;
        pdfDoc.HyphenationOptions.HyphenateCaps = true;
        pdfDoc.HyphenationOptions.ConsecutiveHyphenLimit = 0; // unlimited

        // 6. Render the first page of the PDF to an image.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            PageSet = new PageSet(0), // first page (zero‑based)
            Resolution = 300
        };
        pdfDoc.Save(imagePath, imageOptions);

        // Verify that the image was created.
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Failed to render the PDF to an image.");

        // Optional cleanup (commented out to keep output files for inspection).
        // File.Delete(dictionaryPath);
        // File.Delete(pdfPath);
        // File.Delete(imagePath);
    }
}
