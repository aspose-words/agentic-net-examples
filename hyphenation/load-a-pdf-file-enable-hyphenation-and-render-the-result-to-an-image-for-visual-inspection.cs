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
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary file.
        string dictionaryPath = "hyph_en_US.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for English (United States).
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF output was not created.");

        // Render the document (with hyphenation) to an image.
        string imagePath = "rendered.png";
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // Render the first page.
        };
        doc.Save(imagePath, options);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Image output was not created.");
    }
}
