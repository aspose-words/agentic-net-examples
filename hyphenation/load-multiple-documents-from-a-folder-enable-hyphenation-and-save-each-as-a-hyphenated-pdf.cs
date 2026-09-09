using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationBatchExample
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a few sample DOCX files that will later be hyphenated.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"));

        // Process each document: enable hyphenation and save as PDF.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Apply the locale to all runs in the document so the registered dictionary is used.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                run.Font.LocaleId = new CultureInfo("en-US").LCID;
            }

            // Save as PDF with the same base name.
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }
    }

    // Helper method to create a simple document containing long words that can be hyphenated.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale to English (US) so the registered dictionary applies.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write a paragraph with words that have hyphenation points defined in the dictionary.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        doc.Save(filePath);
    }
}
