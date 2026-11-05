using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create minimal hyphenation dictionaries for English and German.
        const string englishDictPath = "hyph_en_US.dic";
        const string germanDictPath = "hyph_de_DE.dic";

        File.WriteAllText(
            englishDictPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ncommunication=com-mu-ni-ca-tion\n");

        File.WriteAllText(
            germanDictPath,
            "UTF-8\nunabhaengigkeitserklaerung=un-ab-haa-ngi-keit-ser-klä-rung\nkommunikation=kom-mu-ni-ka-tion\n");

        // Register the dictionaries.
        Hyphenation.RegisterDictionary("en-US", englishDictPath);
        Hyphenation.RegisterDictionary("de-DE", germanDictPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow page width to force hyphenation.
        Section firstSection = doc.FirstSection;
        firstSection.PageSetup.PageWidth = 200; // points
        firstSection.PageSetup.LeftMargin = 20;
        firstSection.PageSetup.RightMargin = 20;

        // Set English locale for the first section.
        builder.Font.LocaleId = (int)CultureInfo.GetCultureInfo("en-US").LCID;
        builder.Writeln(
            "extraordinarycharacteristically communication demonstration in a very narrow column.");

        // Insert a new section for German text.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        Section germanSection = doc.Sections[1];
        germanSection.PageSetup.PageWidth = 200;
        germanSection.PageSetup.LeftMargin = 20;
        germanSection.PageSetup.RightMargin = 20;

        // Set German locale for the second section.
        builder.Font.LocaleId = (int)CultureInfo.GetCultureInfo("de-DE").LCID;
        builder.Writeln(
            "unabhaengigkeitserklaerung kommunikation beispiel in einer sehr schmalen spalte.");

        // Save the document as PDF to visualize hyphenation.
        const string outputPath = "HyphenatedOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException(
                $"The expected output file '{outputPath}' was not created.");
        }

        // Clean up temporary dictionary files.
        File.Delete(englishDictPath);
        File.Delete(germanDictPath);
    }
}
