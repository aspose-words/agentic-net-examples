using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a deterministic hyphenation dictionary file.
        const string dictFileName = "hyph_en_US.dic";

        // Minimal dictionary content in OpenOffice format.
        // First line must specify the encoding, subsequent lines contain hyphenation patterns.
        string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary file to the local folder.
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "en-US" locale.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify registration.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with long words that can be hyphenated.
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document as PDF to demonstrate hyphenation.
        const string outputFile = "hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"The expected output file '{outputFile}' was not created.");

        // Optional cleanup (commented out to keep artifacts for inspection).
        // File.Delete(dictFileName);
        // File.Delete(outputFile);
    }
}
