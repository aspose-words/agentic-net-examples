using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "hyphenated.pdf");
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_ru_RU.dic");

        // -----------------------------------------------------------------
        // 1. Create a minimal Russian hyphenation dictionary.
        //    Format: first line is the encoding, subsequent lines are word=pattern.
        // -----------------------------------------------------------------
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "программирование=про-грам-ми-ро-ва-ние\n" +
            "интернационализация=ин-тер-на-цио-на-ли-за-ция\n" +
            "автоматизация=ав-то-мат-из-а-ция\n");

        // Register the dictionary for the Russian locale.
        Aspose.Words.Hyphenation.RegisterDictionary("ru-RU", dictPath);

        // -----------------------------------------------------------------
        // 2. Build a sample document containing Russian text.
        //    Use a narrow page width so that words need to wrap and hyphenate.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
        builder.Writeln("автоматизация программирование интернационализация");

        // Narrow the page to force line breaks.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as DOCX (demonstrates the load‑save cycle).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the saved DOCX and verify the dictionary is still registered.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("ru-RU"))
            throw new InvalidOperationException("Russian hyphenation dictionary was not registered.");

        // Save the loaded document as PDF; hyphenation is applied during layout.
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 4. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF output was not created.");

        Console.WriteLine("PDF created successfully at: " + pdfPath);
    }
}
