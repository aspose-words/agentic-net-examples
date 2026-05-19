using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with custom line spacing (Exactly 30 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
        builder.ParagraphFormat.LineSpacing = 30;
        builder.Writeln("Paragraph with exactly 30 points line spacing.");

        // Add a paragraph with custom line spacing (AtLeast 20 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
        builder.ParagraphFormat.LineSpacing = 20;
        builder.Writeln("Paragraph with at least 20 points line spacing.");

        // Add a paragraph that uses the default line spacing.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 0; // 0 means default multiple (12 points).
        builder.Writeln("Paragraph with default line spacing.");

        // Reset line spacing for all paragraphs to the default.
        // The default is a Multiple rule with a line spacing of 0 (which resolves to 12 points).
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            para.ParagraphFormat.LineSpacing = 0;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResetLineSpacing.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
