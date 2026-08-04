using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Folder that will contain the sample DOCX files.
        const string inputFolder = "InputDocs";
        // Folder where the resulting PDFs will be saved.
        const string outputFolder = "OutputPdfs";

        // Ensure clean environment.
        PrepareFolder(inputFolder);
        PrepareFolder(outputFolder);

        // Create minimal hyphenation dictionaries for English (US) and German (Switzerland).
        CreateHyphenationDictionary("en-US", "hyph_en_US.dic",
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");
        CreateHyphenationDictionary("de-CH", "hyph_de_CH.dic",
            "aussergewöhnlich=aus-ser-gewö-öhnlich\n" +
            "internationalisierung=in-ter-na-tion-a-li-sie-rung\n" +
            "kommunikation=ko-mmu-ni-ka-tion\n");

        // Register the dictionaries so that Aspose.Words can use them during layout.
        Hyphenation.RegisterDictionary("en-US", "hyph_en_US.dic");
        Hyphenation.RegisterDictionary("de-CH", "hyph_de_CH.dic");

        // Create two sample DOCX files – one English, one German.
        CreateSampleDocument(Path.Combine(inputFolder, "EnglishSample.docx"),
            "en-US",
            "extraordinarycharacteristically internationalization communication extraordinarycharacteristically internationalization communication");

        CreateSampleDocument(Path.Combine(inputFolder, "GermanSample.docx"),
            "de-CH",
            "aussergewöhnlich internationalisierung kommunikation aussergewöhnlich internationalisierung kommunikation");

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docxPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            // Optional: fine‑tune hyphenation behaviour.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

            // Determine the language from the file name (simple heuristic).
            string language = docxPath.Contains("English") ? "en-US" : "de-CH";

            // Apply the language to all runs in the document.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                run.Font.LocaleId = new CultureInfo(language).LCID;
            }

            // Build the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        Console.WriteLine("Batch hyphenation and PDF conversion completed successfully.");
    }

    // Ensures that a folder exists and is empty.
    private static void PrepareFolder(string folderPath)
    {
        if (Directory.Exists(folderPath))
            Directory.Delete(folderPath, true);
        Directory.CreateDirectory(folderPath);
    }

    // Writes a hyphenation dictionary file with the supplied content.
    private static void CreateHyphenationDictionary(string language, string fileName, string content)
    {
        // The first line must be the encoding identifier.
        string fullContent = "UTF-8\n" + content;
        File.WriteAllText(fileName, fullContent);
    }

    // Creates a simple DOCX file containing the supplied text and sets a narrow page width
    // to force line wrapping (and thus hyphenation) to be visible.
    private static void CreateSampleDocument(string filePath, string language, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to increase the chance of hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write the text.
        builder.Writeln(text);

        // Apply the language to the runs.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = new CultureInfo(language).LCID;
        }

        // Save the DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
