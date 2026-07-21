using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // File names for the temporary artifacts.
        const string docPath = "sample.docx";
        const string dicPath = "hyph_ru_RU.dic";
        const string pdfPath = "hyphenated_ru_RU.pdf";

        // 1. Create a minimal Russian hyphenation dictionary.
        // The first line must be "UTF-8". Subsequent lines contain word=hyphenated‑pieces.
        File.WriteAllText(dicPath,
            "UTF-8\n" +
            "автомобилизация=авто-мо-били-за-ци-я\n" +
            "интернационализация=интер-на-цио-на-ли-за-ци-я\n");

        // 2. Register the dictionary for the "ru-RU" language.
        Aspose.Words.Hyphenation.RegisterDictionary("ru-RU", dicPath);
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("ru-RU"))
            throw new InvalidOperationException("Failed to register the Russian hyphenation dictionary.");

        // 3. Build a sample document containing Russian text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the font locale to Russian so that the ru‑RU hyphenation dictionary is used.
        builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
        builder.Font.Size = 12;

        // Write a paragraph with words that can be hyphenated.
        builder.Writeln("автомобилизация интернационализация");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the DOCX.
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // 4. Load the DOCX and save it as PDF.
        Document loaded = new Document(docPath);
        loaded.HyphenationOptions.AutoHyphenation = true; // reaffirm
        loaded.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup (commented out).
        // File.Delete(docPath);
        // File.Delete(dicPath);
        // File.Delete(pdfPath);
    }
}
