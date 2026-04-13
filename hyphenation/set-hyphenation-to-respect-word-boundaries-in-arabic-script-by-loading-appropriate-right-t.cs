using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a minimal Arabic hyphenation dictionary file.
        // -----------------------------------------------------------------
        string dictFilePath = Path.Combine(artifactsDir, "hyph_ar.dic");
        string dictContent =
@"SET UTF-8
LEFTHYPHENMIN 2
RIGHTHYPHENMIN 2
% No patterns – this minimal file is sufficient for registration.";
        File.WriteAllText(dictFilePath, dictContent);

        // -----------------------------------------------------------------
        // 2. Register the Arabic dictionary with Aspose.Words.
        // -----------------------------------------------------------------
        using (FileStream dictStream = File.OpenRead(dictFilePath))
        {
            Hyphenation.RegisterDictionary("ar", dictStream);
        }

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("ar"))
            throw new InvalidOperationException("Arabic hyphenation dictionary registration failed.");

        // -----------------------------------------------------------------
        // 3. Build a document containing Arabic text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points (~2.8 inches)
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Set the font locale to Arabic (Saudi Arabia) so that the Arabic
        // hyphenation dictionary is used.
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;
        builder.Font.Size = 12;

        // Arabic sample text long enough to require hyphenation.
        string arabicText = "هذا نص عربي طويل يحتاج إلى تقسيم الكلمات عبر السطر لتوضيح عملية الفواصل في المستند.";
        builder.Writeln(arabicText);

        // -----------------------------------------------------------------
        // 4. Enable automatic hyphenation for the document.
        // -----------------------------------------------------------------
        doc.HyphenationOptions.AutoHyphenation = true;

        // -----------------------------------------------------------------
        // 5. Save the document as PDF.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "ArabicHyphenation.pdf");
        doc.Save(outputPath);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output PDF was not created.", outputPath);
    }
}
