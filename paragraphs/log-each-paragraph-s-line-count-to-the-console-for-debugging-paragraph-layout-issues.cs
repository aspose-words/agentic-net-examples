using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs with varying content.
        builder.Writeln("First paragraph.\nIt has two lines.");
        builder.Writeln("Second paragraph with a single line.");
        builder.Writeln("Third paragraph.\nLine two.\nLine three.");

        // Save the document (optional, demonstrates lifecycle compliance).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleOutput.docx");
        doc.Save(outputPath);

        // Iterate through each paragraph and log an approximate line count.
        int paragraphIndex = 1;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Approximate line count by counting newline characters in the paragraph text.
            string text = paragraph.GetText(); // Includes paragraph break at the end.
            int lineCount = 1; // At least one line exists.
            foreach (char c in text)
            {
                if (c == '\n')
                    lineCount++;
            }

            Console.WriteLine($"Paragraph {paragraphIndex}: Approximate line count = {lineCount}");
            paragraphIndex++;
        }
    }
}
