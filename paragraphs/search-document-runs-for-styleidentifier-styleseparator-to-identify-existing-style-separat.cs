using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the line with Heading1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This is heading text. ");

        // Insert a style separator so the next text can have a different style on the same line.
        builder.InsertStyleSeparator();

        // Second part of the line with Quote style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Write("This is a quote.");

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document (optional, just to visualize the result).
        string docPath = Path.Combine(outputDir, "StyleSeparatorExample.docx");
        doc.Save(docPath);

        // Search for paragraphs that contain a style separator.
        Console.WriteLine("Paragraphs that contain a style separator:");
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int index = 0;
        foreach (Paragraph para in paragraphs)
        {
            if (para.BreakIsStyleSeparator)
            {
                Console.WriteLine($"- Paragraph index {index} (text starts with: \"{para.GetText().Trim()}\" )");
                // Example processing: change the style of the following paragraph part.
                // Here we simply output that a style separator was found.
            }
            index++;
        }

        // The program finishes without waiting for user input.
    }
}
