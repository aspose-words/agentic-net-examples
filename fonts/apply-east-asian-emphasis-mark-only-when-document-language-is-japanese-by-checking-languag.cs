using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "EastAsianEmphasis.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting and formatting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a common font for the document.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;

        // Define the LCID for Japanese language.
        int japaneseLcid = new CultureInfo("ja-JP", false).LCID;

        // Apply Japanese locale to the builder's font.
        builder.Font.LocaleId = japaneseLcid;

        // Check if the current font language is Japanese before applying an emphasis mark.
        if (builder.Font.LocaleId == japaneseLcid)
        {
            // Apply an East Asian emphasis mark (solid circle above the text).
            builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        }

        // Write Japanese text that will display with the emphasis mark.
        builder.Writeln("強調された日本語テキスト");

        // Clear formatting to reset emphasis and locale for the next run.
        builder.Font.ClearFormatting();

        // Set English locale (no emphasis will be applied).
        builder.Font.LocaleId = new CultureInfo("en-US", false).LCID;

        // Write English text; emphasis mark will remain None.
        builder.Writeln("Regular English text without emphasis");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully at: " + outputPath);
        }
    }
}
