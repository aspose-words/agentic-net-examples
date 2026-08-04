using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationItalianTest
{
    public static void Main()
    {
        // Create a minimal Italian hyphenation dictionary file.
        const string dictFileName = "hyph_it_IT.dic";
        string dictContent =
            "UTF-8\n" +
            "casa=ca-sa\n" +
            "automobile=au-to-mo-bile\n" +
            "straordinariamente=stra-or-di-nar-io-men-te\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the Italian locale.
        Hyphenation.RegisterDictionary("it-IT", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("it-IT"))
            throw new InvalidOperationException("Italian hyphenation dictionary was not registered.");

        // Create a new document and configure page layout to force line breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set the paragraph locale to Italian.
        builder.Font.LocaleId = new CultureInfo("it-IT").LCID;
        builder.Font.Size = 12;

        // Write a paragraph containing words that have hyphenation patterns.
        builder.Writeln("casa automobile straordinariamente");

        // Save the document to PDF – this forces layout and hyphenation.
        const string outputFile = "HyphenatedItalian.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected PDF output was not created.");

        // Optional clean‑up (commented out to keep files for inspection).
        // File.Delete(dictFileName);
        // File.Delete(outputFile);
    }
}
