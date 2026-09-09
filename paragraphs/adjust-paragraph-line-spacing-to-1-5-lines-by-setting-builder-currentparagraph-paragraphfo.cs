using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some sample text.
        builder.Writeln("This paragraph will have 1.5 line spacing.");

        // Set the line spacing rule to Multiple and specify 1.5 lines.
        builder.CurrentParagraph.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "ParagraphLineSpacing.docx"));
    }
}
