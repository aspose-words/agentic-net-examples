using System;
using System.IO;
using Aspose.Words;

public class LineSpacingExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Set the line spacing rule to Multiple and the spacing to 1.5 lines.
        paragraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        paragraph.ParagraphFormat.LineSpacing = 1.5;

        // Validate that the line spacing was set correctly.
        if (paragraph.ParagraphFormat.LineSpacingRule != LineSpacingRule.Multiple ||
            Math.Abs(paragraph.ParagraphFormat.LineSpacing - 1.5) > 0.001)
        {
            throw new InvalidOperationException("Line spacing was not set correctly.");
        }

        // Save the document to disk.
        string outputPath = "LineSpacingExample.docx";
        doc.Save(outputPath);

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }

        // Optionally, write a confirmation to the console (no user interaction required).
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
