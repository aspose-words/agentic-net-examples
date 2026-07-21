using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputDir, "sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "sample2.docx"));

        // Process each document: enable hyphenation and save as PDF.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            // Use a valid hyphenation zone (default is 360 = 0.25 inch).
            doc.HyphenationOptions.HyphenationZone = 360;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 0; // No limit.

            // Narrow page width to force line breaks where hyphenation can occur.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF was not created: {pdfPath}");
        }
    }

    // Helper to create a simple document with long words that can be hyphenated.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Use the same narrow page setup as in processing to keep layout consistent.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        doc.Save(filePath, SaveFormat.Docx);

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create source document: {filePath}");
    }
}
