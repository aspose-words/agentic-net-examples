using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the hyphenation dictionary and the resulting files.
        string dictPath = Path.Combine(outputDir, "hyph_de_CH.dic");
        string docPath = Path.Combine(outputDir, "GermanHyphenation.docx");
        string pdfPath = Path.Combine(outputDir, "GermanHyphenation.pdf");

        // Create a minimal German hyphenation dictionary (OpenOffice .dic format).
        // The first line is the number of patterns, followed by simple patterns.
        // These patterns allow hyphenation after German vowels and demonstrate the feature.
        string[] dictLines =
        {
            "5",   // number of patterns
            "1a",
            "1e",
            "1i",
            "1o",
            "1u"
        };
        File.WriteAllLines(dictPath, dictLines);

        // Register the dictionary for the German (Switzerland) locale.
        using (FileStream dictStream = File.OpenRead(dictPath))
        {
            Hyphenation.RegisterDictionary("de-CH", dictStream);
        }

        // Create a new document and set the locale to German (de-CH).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;

        // Add German text containing a long compound word.
        builder.Writeln(
            "Die Donaudampfschifffahrtsgesellschaft ist ein sehr langes deutsches Wort, das hypheniert werden soll.");

        // Enable automatic hyphenation and configure options.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Use a positive hyphenation zone (default is 360 = 0.25 inch).
        doc.HyphenationOptions.HyphenationZone = 360;
        // Optional: limit consecutive hyphenated lines.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Save the document to PDF (hyphenation is applied during layout).
        doc.Save(pdfPath);
        if (!File.Exists(pdfPath))
            throw new Exception("PDF output was not created.");

        // Also save as DOCX for reference.
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception("DOCX output was not created.");
    }
}
