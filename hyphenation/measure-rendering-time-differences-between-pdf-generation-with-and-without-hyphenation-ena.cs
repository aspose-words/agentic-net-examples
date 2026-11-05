using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping where hyphenation can occur.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary locally.
        string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(
            dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the en-US locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // -------------------- Measure without hyphenation --------------------
        doc.HyphenationOptions.AutoHyphenation = false;
        Stopwatch swWithout = Stopwatch.StartNew();
        doc.Save("without_hyphenation.pdf", SaveFormat.Pdf);
        swWithout.Stop();

        // -------------------- Measure with hyphenation --------------------
        doc.HyphenationOptions.AutoHyphenation = true;
        Stopwatch swWith = Stopwatch.StartNew();
        doc.Save("with_hyphenation.pdf", SaveFormat.Pdf);
        swWith.Stop();

        // Validate that the PDF files were created.
        if (!File.Exists("without_hyphenation.pdf"))
            throw new InvalidOperationException("PDF without hyphenation was not created.");
        if (!File.Exists("with_hyphenation.pdf"))
            throw new InvalidOperationException("PDF with hyphenation was not created.");

        // Output timing results.
        Console.WriteLine($"PDF generation without hyphenation: {swWithout.ElapsedMilliseconds} ms");
        Console.WriteLine($"PDF generation with hyphenation: {swWith.ElapsedMilliseconds} ms");
    }
}
