using System;
using System.IO;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationBatchConverter
{
    // Entry point of the console application.
    public static void Main()
    {
        // Define folders for input DOCX files, output PDFs and the log file.
        const string inputFolder = "InputDocs";
        const string outputFolder = "OutputPdfs";
        const string logFilePath = "conversion_log.txt";

        // Ensure clean environment.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        if (File.Exists(logFilePath))
            File.Delete(logFilePath);

        // Create a minimal hyphenation dictionary for English (US).
        const string dictionaryFileName = "hyph_en_US.dic";
        CreateHyphenationDictionary(dictionaryFileName);

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryFileName);

        // Seed a sample DOCX file if the input folder is empty.
        SeedSampleDocumentIfNeeded(inputFolder);

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            try
            {
                // Load the source document.
                Document doc = new Document(docxPath);

                // Enable automatic hyphenation.
                doc.HyphenationOptions.AutoHyphenation = true;

                // Ensure the document language matches the registered dictionary.
                // Here we set the locale of all paragraphs to en-US.
                foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    paragraph.ParagraphFormat.Style.Font.LocaleId = new CultureInfo("en-US").LCID;
                }

                // Determine output PDF path.
                string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save as PDF.
                doc.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF was created.
                if (!File.Exists(pdfPath))
                    throw new InvalidOperationException($"PDF was not created for '{docxPath}'.");
            }
            catch (Exception ex)
            {
                // Log any failure to the log file.
                string logEntry = $"Failed to convert '{Path.GetFileName(docxPath)}': {ex.Message}";
                File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
        }

        // Final status output.
        Console.WriteLine("Batch conversion completed.");
        if (File.Exists(logFilePath))
        {
            Console.WriteLine("Some files failed to convert. See log for details:");
            Console.WriteLine(File.ReadAllText(logFilePath));
        }
        else
        {
            Console.WriteLine("All files converted successfully.");
        }
    }

    // Creates a simple hyphenation dictionary file in OpenOffice format.
    private static void CreateHyphenationDictionary(string fileName)
    {
        // The first line must be the encoding, e.g., "UTF-8".
        // Subsequent lines contain word=hyphenation-patterns.
        string[] lines =
        {
            "UTF-8",
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly",
            "internationalization=in-ter-na-tion-al-i-za-tion",
            "communication=com-mu-ni-ca-tion",
            "demonstration=dem-on-stra-tion",
            "hyphenation=hy-phen-a-tion"
        };
        File.WriteAllLines(fileName, lines);
    }

    // Generates a sample DOCX document with long words to trigger hyphenation.
    private static void SeedSampleDocumentIfNeeded(string folder)
    {
        // If the folder already contains a DOCX, do nothing.
        if (Directory.GetFiles(folder, "*.docx").Length > 0)
            return;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph with words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication demonstration hyphenation " +
            "extraordinarycharacteristically internationalization communication demonstration hyphenation");

        // Save the sample DOCX.
        string samplePath = Path.Combine(folder, "SampleDocument.docx");
        doc.Save(samplePath, SaveFormat.Docx);
    }
}
