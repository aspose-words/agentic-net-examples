using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationBatchProcessor
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create minimal hyphenation dictionaries.
        CreateDictionary("hyph_en_US.dic", new[]
        {
            "UTF-8",
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly",
            "communication=com-mu-ni-ca-tion",
            "internationalization=in-ter-na-tion-al-i-za-tion"
        });

        CreateDictionary("hyph_de_CH.dic", new[]
        {
            "UTF-8",
            "aussergewoehnlich=aus-ser-ge-woehn-lich",
            "kommunikation=ko-mmu-ni-ka-tion",
            "internationalisierung=in-ter-na-tion-a-li-sie-rung"
        });

        // Create sample DOCX files for English and German.
        CreateSampleDocument(Path.Combine(inputDir, "doc_en.docx"), "en-US",
            "extraordinarycharacteristically internationalization communication extraordinarycharacteristically internationalization communication");

        CreateSampleDocument(Path.Combine(inputDir, "doc_de.docx"), "de-CH",
            "aussergewoehnlich internationalisierung kommunikation aussergewoehnlich internationalisierung kommunikation");

        // Register dictionaries for the languages we will use.
        Hyphenation.RegisterDictionary("en-US", "hyph_en_US.dic");
        Hyphenation.RegisterDictionary("de-CH", "hyph_de_CH.dic");

        // Process each DOCX file in the input folder.
        foreach (string docPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.HyphenateCaps = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.

            // Narrow the page width to force line breaks where hyphenation can occur.
            doc.FirstSection.PageSetup.PageWidth = 300; // Points (~4.2 inches).
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }
    }

    // Helper to write a hyphenation dictionary file.
    private static void CreateDictionary(string fileName, string[] lines)
    {
        File.WriteAllLines(fileName, lines);
        if (!File.Exists(fileName))
            throw new InvalidOperationException($"Dictionary file not created: {fileName}");
    }

    // Helper to create a sample DOCX with specified language and text.
    private static void CreateSampleDocument(string filePath, string cultureName, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale for the paragraph runs.
        builder.Font.LocaleId = new CultureInfo(cultureName).LCID;

        // Write the provided text.
        builder.Writeln(text);

        // Save the document.
        doc.Save(filePath, SaveFormat.Docx);

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Sample document not created: {filePath}");
    }
}
