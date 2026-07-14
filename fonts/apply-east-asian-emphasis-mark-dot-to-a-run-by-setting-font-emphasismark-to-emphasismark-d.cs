using System;
using System.IO;
using Aspose.Words;

public class SetEmphasisMarkExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an East Asian emphasis mark to the current font.
        // The EmphasisMark enumeration does not contain a "Dot" value;
        // the closest supported value is OverSolidCircle.
        builder.Font.EmphasisMark = EmphasisMark.OverSolidCircle;

        // Verify that the emphasis mark was set correctly.
        if (builder.Font.EmphasisMark != EmphasisMark.OverSolidCircle)
            throw new InvalidOperationException("Failed to set EmphasisMark to OverSolidCircle.");

        // Write text that will display the emphasis mark.
        builder.Write("Emphasis text");
        builder.Writeln();

        // Clear formatting to remove the emphasis mark for subsequent text.
        builder.Font.ClearFormatting();

        // Write normal text without emphasis.
        builder.Write("Simple text");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SetEmphasisMark.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
