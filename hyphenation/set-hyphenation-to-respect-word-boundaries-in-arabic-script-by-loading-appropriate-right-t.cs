using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set Arabic locale and right‑to‑left paragraph direction.
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;
        builder.Font.Size = 24;
        builder.ParagraphFormat.Bidi = true;

        // Add enough Arabic text to trigger line wrapping.
        string arabicText = "مثال على اختبار تجزئة الكلمات في النص العربي لتوضيح كيفية عمل الفواصل " +
                            "في المستند عندما يكون عرض الصفحة ضيقاً جداً بحيث تحتاج الكلمات إلى تقطيع.";
        builder.Writeln(arabicText);
        builder.Writeln(arabicText); // repeat to ensure wrapping

        // Narrow the page width to force hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal Arabic hyphenation dictionary.
        const string dictFileName = "hyph_ar_SA.dic";
        string dictContent =
            "UTF-8\n" +
            "مثال=مث-ال\n" +
            "تجزئة=تج-زي-ء\n" +
            "كلمات=كلم-ات\n" +
            "النص=الن-ص\n" +
            "العربي=الع-رب-ي\n" +
            "الفواصل=الف-وا-صل\n" +
            "المستند=الم-ست-ند\n" +
            "ضيقاً=ض-يق-اً\n" +
            "تقطيع=تق-طي-ع";

        File.WriteAllText(dictFileName, dictContent);

        // Register the Arabic hyphenation dictionary.
        Aspose.Words.Hyphenation.RegisterDictionary("ar-SA", dictFileName);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF.
        const string outputPath = "HyphenatedArabic.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF file was not created.");
    }
}
