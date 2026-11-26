using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set line spacing to 150% (multiple of 1.5).
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 1.5;

        // Insert a paragraph with some text.
        builder.Writeln("This paragraph has a custom line height of 150 percent.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Reload the document to verify the line spacing.
        Document loadedDoc = new Document(outputPath);
        Paragraph paragraph = (Paragraph)loadedDoc.GetChild(NodeType.Paragraph, 0, true);
        bool lineSpacingCorrect = paragraph.ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple &&
                                  Math.Abs(paragraph.ParagraphFormat.LineSpacing - 1.5) < 0.001;

        // Output verification result.
        Console.WriteLine("Line spacing set correctly: " + lineSpacingCorrect);
    }
}
