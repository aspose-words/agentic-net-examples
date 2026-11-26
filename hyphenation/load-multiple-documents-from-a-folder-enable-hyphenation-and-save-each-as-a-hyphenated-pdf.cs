using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationBatchExample
{
    public static void Main()
    {
        // Create a temporary working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(baseDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
@"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
demonstration=dem-on-stra-tion
");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create sample DOCX files.
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Add text containing words that can be hyphenated.
            builder.Writeln("extraordinarycharacteristically internationalization communication demonstration");
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create sample document: {docPath}");
        }

        // Process each document: enable hyphenation (already active via registration) and save as PDF.
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        if (docFiles.Length == 0)
            throw new InvalidOperationException("No input documents were found.");

        foreach (string docFile in docFiles)
        {
            Document loadedDoc = new Document(docFile);
            // Ensure the document uses the same locale as the dictionary.
            loadedDoc.Styles.DefaultFont.LocaleId = (int)CultureInfo.GetCultureInfo("en-US").LCID;

            string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            loadedDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF was not created for {docFile}");
        }

        // Verify that PDFs were created.
        int pdfCount = Directory.GetFiles(outputDir, "*.pdf").Length;
        if (pdfCount != docFiles.Length)
            throw new InvalidOperationException("Mismatch between input documents and generated PDFs.");

        // Cleanup (optional): comment out if you want to inspect the files.
        // Directory.Delete(baseDir, true);
    }
}
