using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the LCID for Japanese language.
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;

        // Set font properties for Japanese text.
        builder.Font.Name = "Arial";
        builder.Font.LocaleId = japaneseLcid;

        // Apply an emphasis mark only when the language is Japanese.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write Japanese text with the emphasis mark.
        builder.Writeln("日本語のテキストに強調マークが適用されています。");

        // Clear formatting to reset locale and emphasis.
        builder.Font.ClearFormatting();

        // Set font properties for English text (non-Japanese).
        builder.Font.Name = "Arial";
        builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;

        // No emphasis mark will be applied because the language is not Japanese.
        builder.Writeln("This English text does not have an emphasis mark.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "EastAsianEmphasis.docx");
        doc.Save(outputPath);
    }
}
