using System;
using System.Globalization;
using System.IO;
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

        // Prepare a minimal Arabic hyphenation dictionary.
        string myDir = Path.Combine(AppContext.BaseDirectory, "MyDir");
        Directory.CreateDirectory(myDir);
        string arabicDicPath = Path.Combine(myDir, "hyph_ar_SA.dic");

        // If the dictionary file does not exist, create a simple placeholder.
        if (!File.Exists(arabicDicPath))
        {
            // A minimal .dic file (OpenOffice format) – can be empty or contain a comment.
            File.WriteAllText(arabicDicPath, "# Arabic hyphenation dictionary placeholder");
        }

        // Register the Arabic hyphenation dictionary.
        Hyphenation.RegisterDictionary("ar-SA", arabicDicPath);

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(AppContext.BaseDirectory, "ArtifactsDir");
        Directory.CreateDirectory(artifactsDir);

        // Save the document; hyphenation will be applied during layout.
        doc.Save(Path.Combine(artifactsDir, "ArabicHyphenation.docx"));
    }
}
