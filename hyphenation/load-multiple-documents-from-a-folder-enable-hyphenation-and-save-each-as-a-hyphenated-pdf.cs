using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationBatchExample
{
    public static void Main()
    {
        // Base directories for the example.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "HyphenatedPdfs");

        // Ensure clean folders.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample3.docx"));

        // Process each document: enable hyphenation and save as PDF.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);

            // Set document language to English (US) – the language for which we will hyphenate.
            doc.FirstSection.Body.FirstParagraph.Runs[0].Font.LocaleId = new CultureInfo("en-US").LCID;

            // Configure hyphenation options.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 1440 points)
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenateCaps = true;

            // Narrow the page width to force line wrapping and thus hyphenation.
            doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Save as PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // All PDFs should now exist in the output folder.
        Console.WriteLine($"Hyphenated PDFs have been saved to: {outputDir}");
    }

    // Creates a simple DOCX file with long words that can be hyphenated.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a relatively large font to make hyphenation more visible.
        builder.Font.Size = 24;
        builder.Font.Name = "Times New Roman";

        // Add a paragraph containing long words.
        string longText = "characterization hyperventilation internationalization " +
                          "misinterpretation overcompensation uncharacteristically " +
                          "electroencephalographically";
        builder.Writeln(longText);

        // Save the DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
