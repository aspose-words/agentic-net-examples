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

        // Add some introductory text.
        builder.Writeln("This is an introductory paragraph.");

        // Insert a page break before the first heading and apply Heading 1 style.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Add normal text under the first heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 1.");

        // Insert a page break before the second heading and apply Heading 2 style.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Add normal text under the second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 2.");

        // Insert a page break before the third heading and apply Heading 3 style.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        // Add normal text under the third heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content under heading 3.");

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HeadingsWithPageBreaks.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
