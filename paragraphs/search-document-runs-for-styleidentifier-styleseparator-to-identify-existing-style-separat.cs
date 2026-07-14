using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text with the first style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text is in a Heading style. ");

        // Insert a style separator so the next text can have a different style on the same line.
        builder.InsertStyleSeparator();

        // Write text with a different style after the separator.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Write("This text is in a Quote style.");

        // Save the document to the local file system.
        const string outputPath = "StyleSeparatorExample.docx";
        doc.Save(outputPath);

        // Iterate through all paragraphs and find those whose break is a style separator.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.BreakIsStyleSeparator)
            {
                // Output the paragraph text (trimmed) to the console for demonstration.
                Console.WriteLine("Found style separator paragraph: " + paragraph.GetText().Trim());
            }
        }
    }
}
