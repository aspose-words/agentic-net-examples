using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text with a heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This is a heading. ");

        // Insert a style separator so the next text can have a different style
        // but stay on the same line.
        builder.InsertStyleSeparator();

        // Write text with a different style after the separator.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Write("This is a quote on the same line. ");

        // Save the document so we can inspect it if needed.
        const string outputPath = "StyleSeparatorSearch.docx";
        doc.Save(outputPath);

        // Search all paragraphs for style separators.
        int separatorCount = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.BreakIsStyleSeparator)
                separatorCount++;
        }

        // Report the number of style separators found.
        Console.WriteLine($"Found {separatorCount} style separator(s) in the document.");
    }
}
