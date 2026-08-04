using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using static Aspose.Words.Hyphenation; // Hyphenation is a static class, not a namespace

public class Program
{
    public static void Main()
    {
        // Define file names for the temporary files used in the example
        const string docPath = "sample.docx";
        const string dictPath = "hyph_ru_RU.dic";
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a new document and add Russian text that is long enough
        //    to require hyphenation when the line is wrapped.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("ru-RU").LCID;
        builder.Writeln(
            "Это пример текста, который будет демонстрировать переносы слов при гипенуации. " +
            "Текст достаточно длинный, чтобы в узком столбце возникла необходимость в переносе.");

        // Enable automatic hyphenation for the whole document
        doc.HyphenationOptions.AutoHyphenation = true;

        // Reduce the page width so that hyphenation can actually be observed
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document to disk – this will be re‑loaded later
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Create a minimal Russian hyphenation dictionary.
        //    The first line must contain the encoding (UTF-8), followed by
        //    word=hyphenation‑points lines.
        // -----------------------------------------------------------------
        string dictionaryContent =
            "UTF-8\n" +
            "пример=при-мер\n" +
            "текста=тек-ста\n" +
            "переносы=пе-ре-но-сы\n" +
            "гипенуации=ги-пе-ну-а-ци-и\n" +
            "длинный=длин-ный\n" +
            "необходимость=нео-бхо-дим-ость\n" +
            "столбце=стол-бце";

        File.WriteAllText(dictPath, dictionaryContent);

        // Register the dictionary for the Russian locale (ru-RU)
        RegisterDictionary("ru-RU", dictPath);

        // -----------------------------------------------------------------
        // 3. Load the previously saved document and export it to PDF.
        //    Hyphenation will be applied during layout because the
        //    dictionary is already registered and AutoHyphenation is true.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created successfully
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
