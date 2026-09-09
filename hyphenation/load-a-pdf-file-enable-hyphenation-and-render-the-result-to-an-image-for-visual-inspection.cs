using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationPdfToImage
{
    public static void Main()
    {
        // Paths for temporary files
        const string pdfPath = "sample.pdf";
        const string imagePath = "sample_page1.jpg";
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US)
        // The dictionary format: first line is "UTF-8", subsequent lines are word=hyphenation-points
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new document with narrow page width to force line wrapping
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Narrow the page to make hyphenation visible
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document as PDF
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF back into a Document object
        Document pdfDoc = new Document(pdfPath);

        // Render the first page of the PDF to an image (JPEG)
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            PageSet = new PageSet(0), // first page (zero‑based)
            Resolution = 300
        };
        pdfDoc.Save(imagePath, imgOptions);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Image file was not created.");

        // Clean up temporary files (optional)
        // File.Delete(dictPath);
        // File.Delete(pdfPath);
        // File.Delete(imagePath);
    }
}
