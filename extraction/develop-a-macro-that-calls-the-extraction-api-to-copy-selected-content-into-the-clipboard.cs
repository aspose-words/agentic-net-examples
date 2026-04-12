using System;
using System.IO;
using Aspose.Words;

public class ExtractionExample
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph to copy.");
        builder.Writeln("Third paragraph.");

        // Locate the second paragraph (index 1) in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 2)
        {
            throw new InvalidOperationException("The document does not contain enough paragraphs for extraction.");
        }

        Paragraph targetParagraph = paragraphs[1] as Paragraph;
        if (targetParagraph == null)
        {
            throw new InvalidOperationException("Failed to cast the target node to a Paragraph.");
        }

        // Extract the text of the selected paragraph.
        string extractedText = targetParagraph.GetText();
        if (string.IsNullOrEmpty(extractedText))
        {
            throw new InvalidOperationException("Extracted text is empty.");
        }

        // Write the extracted text to a local file for verification.
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedText);

        // Indicate successful completion (no interactive input).
        Console.WriteLine($"Extraction complete. Text saved to '{outputPath}'.");
    }
}
