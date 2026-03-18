using System;
using System.IO;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Settings;

class ArabicHyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Mark the paragraph as right‑to‑left.
        builder.ParagraphFormat.Bidi = true;

        // Configure the font for Arabic (right‑to‑left) text.
        builder.Font.NameBi = "Arial";
        builder.Font.LocaleIdBi = new CultureInfo("ar-SA").LCID;
        builder.Font.Bidi = true;

        // Add Arabic text that will require hyphenation.
        builder.Writeln("هذا مثال على نص عربي طويل يحتاج إلى تقسيم الكلمات عند نهاية السطر لتجنب تجاوز الهوامش.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // optional

        // Register an Arabic hyphenation dictionary (OpenOffice .dic format).
        // Replace "MyDir" with the actual directory containing the dictionary file.
        string arabicDicPath = Path.Combine("MyDir", "hyph_ar_SA.dic");
        Hyphenation.RegisterDictionary("ar-SA", arabicDicPath);

        // Save the document; hyphenation will be applied during layout.
        doc.Save(Path.Combine("ArtifactsDir", "ArabicHyphenation.docx"));
    }
}
