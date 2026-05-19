using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "JapaneseEmphasis.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the language of the text to Japanese (LCID 1041).
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;
        builder.Font.LocaleId = japaneseLcid;

        // Apply an East Asian emphasis mark only if the language is Japanese.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write sample text.
        builder.Writeln("これは強調マーク付きのテキストです。");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
