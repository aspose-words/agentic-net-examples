using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph with a formatted placeholder.
        builder.Write("Hello ");
        builder.Font.Bold = true;   // Make the placeholder bold.
        builder.Font.Italic = true; // Make the placeholder italic.
        builder.Write("_Name_");

        // Reset the font formatting for the rest of the paragraph.
        builder.Font.ClearFormatting();

        builder.Writeln("!");

        // Replace the placeholder while preserving its original formatting.
        int replacements = doc.Range.Replace("_Name_", "World");

        // Verify that a replacement was made.
        if (replacements == 0)
            throw new InvalidOperationException("No occurrences were replaced.");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
