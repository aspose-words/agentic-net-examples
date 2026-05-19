using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // File names for the temporary dictionary, source document and output document.
        const string dictionaryPath = "hyph_en_US.dic";
        const string sourceDocPath = "source.docx";
        const string outputDocPath = "hyphenated.docx";

        // Minimal hyphenation dictionary in OpenOffice format.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary to disk.
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary for the "en-US" language.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a source document with long words that can be hyphenated.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Use a large font size to make hyphenation visible.
        builder.Font.Size = 24;

        // Write a paragraph containing the words defined in the dictionary.
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and thus hyphenation.
        sourceDoc.FirstSection.PageSetup.PageWidth = 200; // points
        sourceDoc.FirstSection.PageSetup.LeftMargin = 20;
        sourceDoc.FirstSection.PageSetup.RightMargin = 20;

        // Save the source document.
        sourceDoc.Save(sourceDocPath);

        // Load the previously saved document.
        Document doc = new Document(sourceDocPath);

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 20 = 36 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document with hyphenation applied.
        doc.Save(outputDocPath);

        // Verify that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException("The hyphenated document was not created.");

        // Clean up temporary files (optional). Comment out the following lines if you wish to inspect the files.
        File.Delete(dictionaryPath);
        File.Delete(sourceDocPath);
        // The output file is left on disk for the user to inspect.
    }
}
