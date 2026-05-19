using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names for the temporary dictionary and the resulting PDF.
        const string dicPath = "hyph_ru_RU.dic";
        const string pdfPath = "Hyphenated_ru_RU.pdf";

        // Create a minimal Russian hyphenation dictionary.
        // The first line must be the encoding, followed by word‑hyphenation patterns.
        File.WriteAllText(
            dicPath,
            "UTF-8\n" +
            "программирование=про-грам-ми-ро-ва-ние\n" +
            "интернационализация=ин-тер-на-цио-на-ли-за-ци-я\n");

        // Register the dictionary for the Russian locale.
        Aspose.Words.Hyphenation.RegisterDictionary("ru-RU", dicPath);

        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and thus hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the font locale to Russian and a readable size.
        builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
        builder.Font.Size = 24;

        // Write a paragraph containing words that can be hyphenated.
        builder.Writeln(
            "программирование интернационализация программирование интернационализация");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional cleanup of the temporary dictionary file.
        // File.Delete(dicPath);
    }
}
