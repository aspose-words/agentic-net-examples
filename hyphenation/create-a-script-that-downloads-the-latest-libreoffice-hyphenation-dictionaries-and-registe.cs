using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        const string dictionaryFile = "hyph_en_US.dic";
        const string outputPdf = "Hyphenated.pdf";

        // Minimal hyphenation dictionary (OpenOffice format)
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary to the local file system
        File.WriteAllText(dictionaryFile, dictionaryContent);

        // Register the dictionary for the "en-US" locale
        Hyphenation.RegisterDictionary("en-US", dictionaryFile);

        // Verify that registration succeeded
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a blank document and add sample text
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow page width to force line wrapping
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch (360 twips)

        // Save the document as PDF
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException($"The expected output file '{outputPdf}' was not created.");

        // Optional: clean up the temporary dictionary file
        // File.Delete(dictionaryFile);
    }
}
