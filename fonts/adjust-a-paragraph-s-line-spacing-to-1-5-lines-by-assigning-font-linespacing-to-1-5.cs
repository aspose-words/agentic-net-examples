using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph line spacing to 1.5 lines.
        // In Aspose.Words, 1 line = 12 points. 1.5 lines = 18 points.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // points

        // Add a sample paragraph.
        builder.Writeln("This paragraph uses 1.5 line spacing.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");

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
