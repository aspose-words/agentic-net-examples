using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a minimal hyphenation dictionary for en‑US.
        string dictionaryPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary (Hyphenation is a static class in Aspose.Words).
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Seed sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Narrow page width to force line wrapping.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;

            // Set document language to en‑US.
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;

            // Write sample text containing long words from the dictionary.
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            builder.Writeln("extraordinarycharacteristically internationalization communication");

            // Save the sample DOCX.
            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create sample document: {docPath}");
        }

        // Batch process: apply hyphenation and export each DOCX to PDF.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            Document loadedDoc = new Document(docxFile);

            // The language of the runs is already set (en‑US), so hyphenation will be applied.

            // Save as PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);
            loadedDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF was not created for {docxFile}");
        }

        // Verify that PDFs were created for all input DOCX files.
        string[] pdfFiles = Directory.GetFiles(outputFolder, "*.pdf");
        if (pdfFiles.Length != docxFiles.Length)
            throw new InvalidOperationException("Mismatch between input DOCX files and generated PDFs.");
    }
}
