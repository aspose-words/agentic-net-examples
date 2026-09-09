using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Japanese language identifier.
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;

        // Set the builder's font language to Japanese.
        builder.Font.LocaleId = japaneseLcid;

        // Apply an emphasis mark only if the language is Japanese.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write Japanese text that will display the emphasis mark.
        builder.Writeln("こんにちは世界"); // "Hello World" in Japanese

        // Clear formatting to reset language and emphasis.
        builder.Font.ClearFormatting();

        // Set language to English (no emphasis will be applied).
        builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;
        builder.Writeln("Hello world!");

        // Define output path and ensure the directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "EastAsianEmphasis.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
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
