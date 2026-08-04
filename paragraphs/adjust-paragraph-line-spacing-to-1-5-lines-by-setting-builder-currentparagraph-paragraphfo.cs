using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some sample text.
        builder.Writeln("This paragraph will have 1.5 line spacing applied to it.");
        builder.Writeln("The second line of the same paragraph demonstrates the spacing.");

        // Set the line spacing rule to Multiple (number of lines) and apply 1.5 lines.
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document to the output file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ParagraphLineSpacing.docx");
        doc.Save(outputPath);
    }
}
