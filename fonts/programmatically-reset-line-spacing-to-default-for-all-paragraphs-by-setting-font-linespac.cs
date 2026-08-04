using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs with different formatting
        builder.Writeln("First paragraph with default line spacing.");
        builder.Font.Size = 24;
        builder.Writeln("Second paragraph with larger font size.");
        builder.Font.Size = 12;
        builder.Writeln("Third paragraph with normal font size.");

        // Reset line spacing for all paragraphs to the default values
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Default line spacing is defined by the Multiple rule with a value of 0 (auto)
            paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            paragraph.ParagraphFormat.LineSpacing = 0;
        }

        // Save the document
        string outputPath = Path.Combine(artifactsDir, "ResetLineSpacing.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple verification that the file was created
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully to: {outputPath}"
            : "Failed to save the document.");
    }
}
