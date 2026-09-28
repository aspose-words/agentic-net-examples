using System;
using System.IO;
using Aspose.Words;

public class ExportParagraphsWithLineNumbers
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Export the document to plain text with line numbers prefixed.
        string outputPath = "output.txt";
        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            // Retrieve all paragraph nodes in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            int lineNumber = 1;

            foreach (Paragraph para in paragraphs)
            {
                // Get the paragraph text without the trailing paragraph mark.
                string text = para.GetText().TrimEnd('\r', '\n');
                writer.WriteLine($"{lineNumber}: {text}");
                lineNumber++;
            }
        }

        // Indicate completion.
        Console.WriteLine($"Document saved to {outputPath} with line numbers.");
    }
}
