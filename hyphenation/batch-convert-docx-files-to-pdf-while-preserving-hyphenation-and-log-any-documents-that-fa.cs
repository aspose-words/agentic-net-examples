using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Directories for input DOCX files, output PDFs and the log file.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string logPath = Path.Combine(baseDir, "conversion_log.txt");
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        if (File.Exists(logPath))
            File.Delete(logPath);

        // Create a minimal hyphenation dictionary and register it.
        CreateDictionaryFile(dictPath);
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Generate sample DOCX files that contain words suitable for hyphenation.
        CreateSampleDocuments(inputDir);

        // Batch convert each DOCX to PDF while preserving hyphenation.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            string fileName = Path.GetFileNameWithoutExtension(docxPath);
            string pdfPath = Path.Combine(outputDir, fileName + ".pdf");

            try
            {
                // Load the DOCX document.
                Document doc = new Document(docxPath);

                // Enable automatic hyphenation for the document.
                doc.HyphenationOptions.AutoHyphenation = true;
                doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 pt units).

                // Save as PDF.
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created: {pdfPath}");
            }
            catch (Exception ex)
            {
                // Log any conversion failures.
                File.AppendAllText(logPath, $"Failed to convert '{docxPath}': {ex.Message}{Environment.NewLine}");
            }
        }

        // Optional: report conversion summary.
        int total = Directory.GetFiles(inputDir, "*.docx").Length;
        int succeeded = Directory.GetFiles(outputDir, "*.pdf").Length;
        Console.WriteLine($"Conversion finished. {succeeded}/{total} documents converted successfully.");
        if (File.Exists(logPath))
            Console.WriteLine($"See log for details: {logPath}");
    }

    // Creates a simple hyphenation dictionary file.
    private static void CreateDictionaryFile(string path)
    {
        // Dictionary format: first line "UTF-8", then word=hy-phen-a-tion entries.
        string content = "UTF-8\n" +
                         "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
                         "internationalization=in-ter-na-tion-al-i-za-tion\n" +
                         "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(path, content);
    }

    // Generates a few sample DOCX files with long words that can be hyphenated.
    private static void CreateSampleDocuments(string folder)
    {
        // Create three documents; each contains the same long words to trigger hyphenation.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow the page width so that long words wrap and hyphenate.
            doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Use a reasonably large font to make wrapping more likely.
            builder.Font.Size = 24;
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;

            // Write a paragraph containing words that exist in the dictionary.
            builder.Writeln("extraordinarycharacteristically internationalization communication");

            string filePath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(filePath);
        }
    }
}
