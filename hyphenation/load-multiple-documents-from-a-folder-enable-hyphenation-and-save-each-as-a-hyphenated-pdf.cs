using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
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

        // Register the dictionary once – it will be used for all documents.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"),
            "extraordinarycharacteristically internationalization communication");
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"),
            "communication communication communication communication communication");

        // Process each document in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: adjust page setup to increase chance of hyphenation.
            doc.FirstSection.PageSetup.PageWidth = 200; // points
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Save as PDF.
            string outputFile = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            doc.Save(outputFile, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Failed to create PDF: {outputFile}");
        }
    }

    // Helper method to create a simple DOCX with given text.
    private static void CreateSampleDocument(string fileName, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set language to English (US) so the registered dictionary applies.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln(text);

        // Save the document.
        doc.Save(fileName, SaveFormat.Docx);

        // Verify creation.
        if (!File.Exists(fileName))
            throw new InvalidOperationException($"Failed to create source document: {fileName}");
    }
}
