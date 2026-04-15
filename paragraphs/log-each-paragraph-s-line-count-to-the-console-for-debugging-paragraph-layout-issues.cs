using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several sample paragraphs.
        builder.Writeln("First paragraph with some text.");
        builder.Writeln("Second paragraph with a bit more text that might wrap onto multiple lines depending on page width.");
        builder.Writeln("Third paragraph.\nIt contains an explicit line break.");
        builder.Writeln("Fourth paragraph.");

        // Retrieve all paragraph nodes in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        int paragraphIndex = 1;

        // Log an approximate line count for each paragraph.
        // Here we use the number of runs as a simple proxy for line count.
        foreach (Paragraph para in paragraphs)
        {
            int approximateLineCount = para.Runs.Count;
            Console.WriteLine($"Paragraph {paragraphIndex}: Approximate line count (runs) = {approximateLineCount}");
            paragraphIndex++;
        }

        // Save the document to the output file.
        doc.Save("ParagraphLineCounts.docx");
    }
}
