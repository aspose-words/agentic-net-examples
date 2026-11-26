using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for en-US.
        string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary and enable automatic hyphenation for the document.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // Render the first page of the PDF to an image.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // first page (zero‑based index)
        };
        string imagePath = "rendered.png";
        pdfDoc.Save(imagePath, options);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Rendered image was not created.");
    }
}
