using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationPdfGenerator
{
    public static void Main(string[] args)
    {
        // Determine language code (default to en-US if not supplied)
        string language = args.Length >= 2 ? args[1] : "en-US";

        // Prepare a hyphenation dictionary file for the selected language.
        string dictFileName = $"hyph_{language}.dic";
        // Minimal dictionary content – includes a long word that can be hyphenated.
        string dictContent = "UTF-8\nextraordinarycharacteristically=ex-tra-or-di-na-ry-char-ac-ter-is-ti-cal-ly\n";
        File.WriteAllText(dictFileName, dictContent);
        Hyphenation.RegisterDictionary(language, dictFileName);

        Document doc;
        string outputPath;

        // If a valid input document path is provided, load it; otherwise create a sample document.
        if (args.Length >= 1 && File.Exists(args[0]))
        {
            string inputPath = args[0];
            doc = new Document(inputPath);
            outputPath = Path.Combine(
                Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory(),
                Path.GetFileNameWithoutExtension(inputPath) + "_hyphenated.pdf");
        }
        else
        {
            // Create a sample document with narrow page width to force line wrapping.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Size = 24;
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            // Narrow page width so words wrap and hyphenation can be seen.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            outputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample_hyphenated.pdf");
        }

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Apply the language locale to all runs.
        CultureInfo culture = new CultureInfo(language);
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = culture.LCID;
        }

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The hyphenated PDF was not created.");

        // Clean up the temporary dictionary file.
        try { File.Delete(dictFileName); } catch { /* ignore cleanup errors */ }
    }
}
