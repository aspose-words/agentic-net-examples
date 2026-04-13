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

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the LCID for Japanese language.
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;

        // Set the builder's font locale to Japanese.
        builder.Font.LocaleId = japaneseLcid;

        // Write a normal line of Japanese text.
        builder.Writeln("こんにちは世界"); // "Hello world" in Japanese.

        // Check if the current language is Japanese before applying an emphasis mark.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            // Apply an East Asian emphasis mark (solid circle above the text).
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write a line that will have the emphasis mark applied.
        builder.Writeln("強調テキスト"); // "Emphasis text" in Japanese.

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

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
