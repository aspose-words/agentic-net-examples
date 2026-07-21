using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long sentence that can be hyphenated.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line breaks and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Create a minimal hyphenation dictionary for en‑US.
        const string dictFileName = "hyph_en_US.dic";
        string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary for en-US was not registered.");

        // Save the document as PDF – hyphenated words will appear in the output.
        const string outputFile = "hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
