using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("This paragraph will have double line spacing.");

        // Set line spacing to double (24 points) using the Multiple rule.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 24; // 2 * 12 points (default line height)

        // Validate that the settings were applied correctly.
        bool ruleIsMultiple = builder.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple;
        bool spacingIsDouble = Math.Abs(builder.ParagraphFormat.LineSpacing - 24) < 0.001;

        Console.WriteLine($"LineSpacingRule set to Multiple: {ruleIsMultiple}");
        Console.WriteLine($"LineSpacing set to double (24 points): {spacingIsDouble}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleLineSpacing.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        Console.WriteLine($"Document saved successfully: {File.Exists(outputPath)}");
    }
}
