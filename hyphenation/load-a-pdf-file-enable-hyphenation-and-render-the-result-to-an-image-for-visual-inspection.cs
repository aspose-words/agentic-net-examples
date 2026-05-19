using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string pdfPath = "hyphenated.pdf";
        const string imagePath = "hyphenated.jpg";

        // 1. Create a sample document with text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Narrow the page to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // 2. Create a minimal hyphenation dictionary for English (US).
        // The dictionary format is the OpenOffice hyphenation format.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // 3. Register the dictionary and enable automatic hyphenation.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        // 4. Save the document as PDF.
        doc.Save(pdfPath);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF output was not created.");

        // 5. Load the PDF back into a Document.
        Document pdfDoc = new Document(pdfPath);
        // Ensure hyphenation options are still enabled after loading.
        pdfDoc.HyphenationOptions.AutoHyphenation = true;

        // 6. Render the first page of the PDF to an image for visual inspection.
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: increase resolution for clearer output.
            Resolution = 300
        };
        pdfDoc.Save(imagePath, imgOptions);
        if (!File.Exists(imagePath))
            throw new InvalidOperationException("Image output was not created.");

        // Cleanup: optional removal of temporary files.
        // File.Delete(dictionaryPath);
        // File.Delete(pdfPath);
    }
}
