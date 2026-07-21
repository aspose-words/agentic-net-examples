using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write Arabic text that is long enough to require hyphenation when the line width is narrow.
        // The word "مثال" (example) will be hyphenated according to the dictionary we provide.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;
        builder.Writeln("هذا نص تجريبي يحتوي على كلمة مثال لتوضيح كيفية تطبيق الفواصل في النص العربي عند الحاجة إلى كسر السطر.");

        // Configure the document to use automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: limit consecutive hyphens and set hyphenation zone.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Make the paragraph right‑to‑left.
        doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi = true;

        // Narrow the page width to force line wrapping and thus hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal Arabic hyphenation dictionary.
        // The format follows the OpenOffice hyphenation dictionary specification.
        const string dictFileName = "hyph_ar_SA.dic";
        string dictContent = "UTF-8\nمثال=مث-ال\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the Arabic (Saudi Arabia) locale.
        Hyphenation.RegisterDictionary("ar-SA", dictFileName);

        // Save the document to PDF so that hyphenation can be visually inspected.
        const string outputFile = "HyphenatedArabic.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"The expected output file '{outputFile}' was not created.");

        // Clean up the temporary dictionary file.
        File.Delete(dictFileName);
    }
}
