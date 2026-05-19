using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Base working directory for the example.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationExample");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");

        // Ensure clean environment.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US).
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create sample DOCX files.
        string[] sampleFileNames = { "Sample1.docx", "Sample2.docx", "Sample3.docx" };
        foreach (string fileName in sampleFileNames)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set a narrow page width to force line wrapping.
            doc.FirstSection.PageSetup.PageWidth = 200; // points
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Use English locale.
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;
            builder.Font.Size = 12;

            // Write text that contains long words suitable for hyphenation.
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            builder.Writeln("extraordinarycharacteristically internationalization communication");

            // Save the source document.
            string inputPath = Path.Combine(inputDir, fileName);
            doc.Save(inputPath);
            if (!File.Exists(inputPath))
                throw new InvalidOperationException($"Failed to create source document: {inputPath}");
        }

        // Process each document: enable hyphenation and save as PDF.
        foreach (string sourcePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(sourcePath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            // Optional: configure additional hyphenation settings.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
            doc.HyphenationOptions.HyphenateCaps = true;

            // Save as PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(sourcePath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate output.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // All PDFs should now exist.
        string[] pdfFiles = Directory.GetFiles(outputDir, "*.pdf");
        if (pdfFiles.Length != sampleFileNames.Length)
            throw new InvalidOperationException("The number of generated PDFs does not match the expected count.");

        // Example completed successfully.
        Console.WriteLine("Hyphenated PDFs have been generated in: " + outputDir);
    }
}
