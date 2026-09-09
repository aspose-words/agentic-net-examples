using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the temporary Italian hyphenation dictionary.
        const string dictFile = "hyph_it_IT.dic";

        // Create a minimal valid dictionary file.
        // The first line must be the encoding identifier.
        // Subsequent lines contain word=hyphenated-pattern.
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-na-ri-a-men-te\n");

        // Register the dictionary for the Italian locale.
        Hyphenation.RegisterDictionary("it-IT", dictFile);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("it-IT"))
            throw new InvalidOperationException("Italian hyphenation dictionary was not registered.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font locale to Italian so that the hyphenation engine uses the correct language.
        builder.Font.LocaleId = new CultureInfo("it-IT").LCID;
        builder.Font.Size = 24;

        // Write a word that can be hyphenated according to the dictionary.
        builder.Writeln("extraordinariamente");

        // Narrow the page width to force line wrapping and trigger hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF (any format works; PDF demonstrates layout).
        const string outputFile = "HyphenatedItalian.pdf";
        doc.Save(outputFile);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The hyphenated PDF was not created.");

        // Clean up temporary dictionary file (optional).
        File.Delete(dictFile);
    }
}
