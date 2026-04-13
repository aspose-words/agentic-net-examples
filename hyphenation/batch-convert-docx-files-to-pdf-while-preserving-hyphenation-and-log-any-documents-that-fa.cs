using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (en-US).
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        CreateEnglishHyphenationDictionary(dictPath);

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary for en-US.");

        // Seed sample DOCX files.
        int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            CreateSampleDocument(docPath);
        }

        // Prepare log file.
        string logPath = Path.Combine(baseDir, "conversion_log.txt");
        using (StreamWriter logWriter = new StreamWriter(logPath, false))
        {
            // Process each DOCX file.
            foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
            {
                try
                {
                    // Load the document.
                    Document doc = new Document(docFile);

                    // Ensure automatic hyphenation is enabled.
                    doc.HyphenationOptions.AutoHyphenation = true;
                    doc.HyphenationOptions.HyphenateCaps = true;
                    doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
                    doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.

                    // Set locale to match the registered dictionary.
                    doc.FirstSection.Body.FirstParagraph.Runs[0].Font.LocaleId = new System.Globalization.CultureInfo("en-US").LCID;

                    // Save as PDF.
                    string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
                    string pdfPath = Path.Combine(outputDir, pdfFileName);
                    doc.Save(pdfPath, SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    // Log failure.
                    logWriter.WriteLine($"Failed to convert '{Path.GetFileName(docFile)}': {ex.Message}");
                }
            }
        }

        // Optional: indicate completion (no interactive input).
        Console.WriteLine("Batch conversion completed.");
    }

    // Creates a very small English hyphenation dictionary in OpenOffice format.
    private static void CreateEnglishHyphenationDictionary(string filePath)
    {
        // The first line is the number of patterns; we provide a few simple patterns.
        // This dictionary is not exhaustive but sufficient for demonstration.
        string[] lines =
        {
            "5",          // Number of patterns.
            "ab1c",       // Example pattern.
            "de2f",
            "ghi3j",
            "kl4m",
            "no5p"
        };
        File.WriteAllLines(filePath, lines);
    }

    // Generates a sample DOCX with long text to trigger hyphenation.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to increase line wrapping.
        builder.PageSetup.PageWidth = 300; // Points (~4.17 inches).

        // Insert a paragraph with repetitive long words.
        builder.Font.Size = 12;
        builder.Writeln("Antidisestablishmentarianism is a long word that often needs hyphenation. " +
                        "Supercalifragilisticexpialidocious also demonstrates hyphenation capabilities. " +
                        "Pseudopseudohypoparathyroidism is another example of a word that may be hyphenated.");

        // Enable hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.

        // Set the language of the text to English (US) to match the dictionary.
        builder.Font.LocaleId = new System.Globalization.CultureInfo("en-US").LCID;

        doc.Save(filePath, SaveFormat.Docx);
    }
}
