using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationPdfToImage
{
    public static void Main()
    {
        // Output file names (created in the program's working directory)
        const string pdfPath = "sample.pdf";
        const string imagePath = "sample.png";
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US)
        // First line is the encoding, subsequent lines are word=hyphenation-points
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document with narrow page width to force line wrapping
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page to increase the chance of hyphenation
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation
        doc.HyphenationOptions.AutoHyphenation = true;

        // HyphenationZone must be a non‑negative value; use the default (360) to allow hyphenation up to the margin
        doc.HyphenationOptions.HyphenationZone = 360;

        // Save the document as PDF
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF back
        Document loadedPdf = new Document(pdfPath);

        // Render the first page to a PNG image
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // first page (zero‑based)
        };
        loadedPdf.Save(imagePath, imgOptions);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Image file was not created.");

        // Optional cleanup (uncomment if you want to delete the temporary files)
        // File.Delete(dictPath);
        // File.Delete(pdfPath);
        // File.Delete(imagePath);
    }
}
