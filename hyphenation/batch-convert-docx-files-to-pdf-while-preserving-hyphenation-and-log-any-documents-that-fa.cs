using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files, output PDFs and the log file.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string logPath = Path.Combine(baseDir, "conversion_log.txt");

        // Clean previous run data.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        if (File.Exists(logPath)) File.Delete(logPath);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (en-US).
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a few sample DOCX files that contain long words to trigger hyphenation.
        for (int i = 1; i <= 3; i++)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping.
            doc.FirstSection.PageSetup.PageWidth = 300;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720;

            // Set locale to en-US so the registered dictionary is used.
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;

            // Write sample text containing words defined in the dictionary.
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            builder.Writeln($"Document number {i} with long words to demonstrate hyphenation.");

            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            doc.Save(docPath);
        }

        // Process each DOCX file: convert to PDF while preserving hyphenation.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            try
            {
                var doc = new Document(docxPath);
                // Ensure hyphenation options are enabled for each loaded document.
                doc.HyphenationOptions.AutoHyphenation = true;
                doc.HyphenationOptions.HyphenateCaps = true;

                string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Validate that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created for '{docxPath}'.");
            }
            catch (Exception ex)
            {
                // Log any failures to the log file.
                File.AppendAllText(logPath, $"Failed to convert '{docxPath}': {ex.Message}{Environment.NewLine}");
            }
        }

        // Final validation: ensure at least one PDF was produced.
        int pdfCount = Directory.GetFiles(outputDir, "*.pdf").Length;
        if (pdfCount == 0)
            throw new InvalidOperationException("No PDF files were generated. Check the log for details.");

        // Optionally, indicate completion (no interactive output required).
    }
}
