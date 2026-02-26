using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class HyphenatePdf
{
    static void Main()
    {
        // Paths to the source PDF and the resulting PDF.
        string inputPath = @"C:\Docs\input.pdf";
        string outputPath = @"C:\Docs\output.pdf";

        // Load the existing PDF document.
        Document doc = new Document(inputPath);

        // Ensure an English hyphenation dictionary is registered.
        // If the dictionary file is not present, the code will throw.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            using (FileStream stream = new FileStream(@"C:\Hyphenation\hyph_en_US.dic", FileMode.Open, FileAccess.Read))
            {
                Hyphenation.RegisterDictionary("en-US", stream);
            }
        }

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: set how close to the right margin hyphenation is allowed (0.5 inch).
        doc.HyphenationOptions.HyphenationZone = 720; // 720 = 0.5 inch (1/20 pt units)
        // Optional: limit consecutive hyphenated lines.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        // Optional: hyphenate words written in all caps.
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document as PDF – hyphenation will be applied during layout.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
