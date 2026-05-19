using System;
using System.IO;
using System.Globalization;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationBatch");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        string logPath = Path.Combine(baseDir, "conversion.log");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        File.WriteAllText(logPath, string.Empty);

        // Create a minimal hyphenation dictionary for English (en-US).
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary once – it will be used for all documents.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create sample DOCX files to demonstrate the batch process.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"),
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"),
            "communication communication communication communication communication " +
            "internationalization internationalization internationalization internationalization");

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            try
            {
                // Load the document.
                Document doc = new Document(docxPath);

                // Enable automatic hyphenation.
                doc.HyphenationOptions.AutoHyphenation = true;

                // Ensure the document language matches the registered dictionary.
                foreach (Run run in doc.GetChildNodes(NodeType.Run, true).OfType<Run>())
                {
                    run.Font.LocaleId = new CultureInfo("en-US").LCID;
                }

                // Narrow the page width to force line wrapping and hyphenation.
                doc.FirstSection.PageSetup.PageWidth = 200;
                doc.FirstSection.PageSetup.LeftMargin = 20;
                doc.FirstSection.PageSetup.RightMargin = 20;

                // Save as PDF.
                string pdfPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException("PDF file was not created.");
            }
            catch (Exception ex)
            {
                // Log any failures.
                File.AppendAllText(logPath,
                    $"{Path.GetFileName(docxPath)}: {ex.Message}{Environment.NewLine}");
            }
        }

        // Optional: write a short summary to the console.
        Console.WriteLine("Batch conversion completed.");
        Console.WriteLine($"PDFs saved to: {outputDir}");
        Console.WriteLine($"Log file: {logPath}");
    }

    // Helper method to create a simple DOCX file with the given text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
