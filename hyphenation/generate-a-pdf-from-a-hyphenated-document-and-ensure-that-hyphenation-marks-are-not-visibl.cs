using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US) in the local folder.
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new document and set a narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add some long words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set a very large hyphenation zone so that hyphens are never placed in the visible area.
        // Using the page width ensures the zone exceeds the printable area.
        doc.HyphenationOptions.HyphenationZone = (int)doc.FirstSection.PageSetup.PageWidth;

        // Save the document as PDF.
        const string pdfFileName = "hyphenated.pdf";
        doc.Save(pdfFileName, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
