using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationBatchProcessor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationBatch");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean folders.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create minimal hyphenation dictionaries.
        string enDictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        string deDictPath = Path.Combine(baseDir, "hyph_de_CH.dic");

        File.WriteAllText(enDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        File.WriteAllText(deDictPath,
            "UTF-8\n" +
            "internationalisierung=in-ter-na-tion-a-lei-tung\n" +
            "kommunikation=kom-mu-ni-ka-tion\n" +
            "extraordinaer=ex-tra-or-di-när\n");

        // Register dictionaries for the required languages.
        Hyphenation.RegisterDictionary("en-US", enDictPath);
        Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // Prepare sample documents.
        var samples = new List<(string FileName, string Language, string Text)>
        {
            ("EnglishSample.docx", "en-US",
                "extraordinarycharacteristically internationalization communication " +
                "extraordinarycharacteristically internationalization communication"),
            ("GermanSample.docx", "de-CH",
                "extraordinaer internationalisierung kommunikation " +
                "extraordinaer internationalisierung kommunikation")
        };

        foreach (var sample in samples)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Set the locale for the paragraph runs.
            builder.Font.LocaleId = new CultureInfo(sample.Language).LCID;
            builder.Font.Size = 12;
            builder.Writeln(sample.Text);

            // Save the source DOCX.
            string docxPath = Path.Combine(inputDir, sample.FileName);
            doc.Save(docxPath);
        }

        // Process each DOCX: apply hyphenation and export to PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docxPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF was not created: {pdfPath}");
        }

        // Optional: indicate completion (no interactive input).
        Console.WriteLine("Batch hyphenation processing completed successfully.");
    }
}
