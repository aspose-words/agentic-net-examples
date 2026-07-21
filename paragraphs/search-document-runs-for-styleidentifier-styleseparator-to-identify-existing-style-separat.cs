using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
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

        // Save the document to the local folder.
        string docPath = "StyleSeparatorDemo.docx";
        doc.Save(docPath);

        // Search for paragraphs that represent a style separator.
        int separatorCount = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.BreakIsStyleSeparator)
            {
                separatorCount++;

                // Example processing: change the font color of the text that follows the separator.
                // The paragraph still contains the runs after the separator, so we can modify them.
                foreach (Run run in paragraph.Runs)
                {
                    // Skip runs that belong to the part before the separator.
                    // In this simple example we just change all runs after the separator.
                    run.Font.Color = System.Drawing.Color.Blue;
                }
            }
        }

        // Output the number of style separators found.
        Console.WriteLine($"Found {separatorCount} style separator(s) in the document.");
    }
}
