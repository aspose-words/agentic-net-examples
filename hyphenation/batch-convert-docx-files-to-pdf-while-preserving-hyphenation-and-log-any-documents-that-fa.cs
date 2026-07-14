using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string logFile = Path.Combine(baseDir, "failed.txt");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        if (File.Exists(logFile)) File.Delete(logFile);

        // Create a minimal hyphenation dictionary for English (en-US)
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create sample DOCX files
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping and hyphenation
            doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Enable automatic hyphenation for the document
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 360; // default

            // Write a paragraph with long words that can be hyphenated
            builder.Font.Size = 12;
            builder.Writeln(
                "extraordinarycharacteristically internationalization communication " +
                "extraordinarycharacteristically internationalization communication");

            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create sample document: {docPath}");
        }

        // Process each DOCX file
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            try
            {
                Document doc = new Document(docxPath);

                // Ensure hyphenation options are set (in case the source lacks them)
                doc.HyphenationOptions.AutoHyphenation = true;
                doc.HyphenationOptions.HyphenateCaps = true;

                string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                string pdfPath = Path.Combine(outputDir, pdfFileName);

                doc.Save(pdfPath, SaveFormat.Pdf);

                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created: {pdfPath}");
            }
            catch (Exception ex)
            {
                // Log failure
                File.AppendAllText(logFile, $"Failed to convert '{docxPath}': {ex.Message}{Environment.NewLine}");
            }
        }

        // Optional: indicate completion (no interactive output required)
        // The program ends here.
    }
}
