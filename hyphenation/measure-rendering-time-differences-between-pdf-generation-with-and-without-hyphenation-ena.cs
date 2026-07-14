using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow page width forces line wrapping, making hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Measure rendering time without hyphenation.
        doc.HyphenationOptions.AutoHyphenation = false;
        const string pdfNoHyphen = "nonhyphenated.pdf";
        Stopwatch swNoHyphen = Stopwatch.StartNew();
        doc.Save(pdfNoHyphen, SaveFormat.Pdf);
        swNoHyphen.Stop();

        // Measure rendering time with hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        const string pdfHyphen = "hyphenated.pdf";
        Stopwatch swHyphen = Stopwatch.StartNew();
        doc.Save(pdfHyphen, SaveFormat.Pdf);
        swHyphen.Stop();

        // Validate that the PDF files were created.
        if (!File.Exists(pdfNoHyphen))
            throw new InvalidOperationException($"File '{pdfNoHyphen}' was not created.");
        if (!File.Exists(pdfHyphen))
            throw new InvalidOperationException($"File '{pdfHyphen}' was not created.");

        // Output the measured times.
        Console.WriteLine($"PDF generation without hyphenation: {swNoHyphen.ElapsedMilliseconds} ms");
        Console.WriteLine($"PDF generation with hyphenation   : {swHyphen.ElapsedMilliseconds} ms");
    }
}
