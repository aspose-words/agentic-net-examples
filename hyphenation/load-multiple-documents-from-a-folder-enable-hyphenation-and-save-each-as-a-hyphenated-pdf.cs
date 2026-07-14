using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US)
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n");

        // Register the dictionary
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create sample source documents
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"),
            "extraordinarycharacteristically internationalization communication");
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"),
            "communication communication communication extraordinarycharacteristically");

        // Process each document: enable hyphenation and save as PDF
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docPath);

            // Narrow page width to force line wrapping and hyphenation
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Enable automatic hyphenation
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 360; // default

            // Save as PDF
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate output
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: clean up dictionary file (not required)
        // File.Delete(dictPath);
    }

    // Helper to create a simple DOCX with given text
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create source document: {filePath}");
    }
}
