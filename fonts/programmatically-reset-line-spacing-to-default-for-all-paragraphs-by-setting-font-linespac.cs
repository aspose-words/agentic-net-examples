using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a custom line spacing (exactly 30 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
        builder.ParagraphFormat.LineSpacing = 30;
        builder.Writeln("Paragraph with custom line spacing (Exactly 30 points).");

        // Add another paragraph with a different custom line spacing (at least 20 points).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;
        builder.ParagraphFormat.LineSpacing = 20;
        builder.Writeln("Another paragraph with custom line spacing (AtLeast 20 points).");

        // Reset line spacing for all paragraphs to the default value.
        // Setting LineSpacing to 0 and using the Multiple rule restores the default behavior.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacing = 0;
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ResetLineSpacing.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
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
