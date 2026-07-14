using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the LCID for Japanese language.
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;

        // Set the builder's font language to Japanese.
        builder.Font.LocaleId = japaneseLcid;

        // Write some Japanese text.
        builder.Writeln("こんにちは世界"); // "Hello World" in Japanese

        // Apply an East Asian emphasis mark only if the current language is Japanese.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write Japanese text that will have the emphasis mark.
        builder.Writeln("強調されたテキスト"); // "Emphasized text"

        // Reset formatting to avoid affecting subsequent text.
        builder.Font.ClearFormatting();

        // Write English text without emphasis.
        builder.Writeln("Hello world!");

        // Save the document to a file.
        string outputPath = "EastAsianEmphasis.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
