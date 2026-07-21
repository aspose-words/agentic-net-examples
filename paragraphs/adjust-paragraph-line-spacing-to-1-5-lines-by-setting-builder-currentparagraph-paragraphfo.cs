using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph of text.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple (optional, ensures the value is treated as a multiple of the default line height).
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;

        // Adjust the line spacing to 1.5 lines.
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Define an output path for the generated document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ParagraphLineSpacing.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
