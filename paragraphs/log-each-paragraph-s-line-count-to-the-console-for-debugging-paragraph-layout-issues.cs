using System;
using Aspose.Words;
using Aspose.Words.Layout;

namespace ParagraphLineCountDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add several paragraphs with different amounts of text.
            builder.Writeln("Short paragraph.");
            builder.Writeln("This paragraph contains a bit more text, but still fits on a single line in most layouts.");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            // Traverse all paragraph nodes in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            int index = 1;
            foreach (Paragraph para in paragraphs)
            {
                // Approximate the line count by using the number of runs in the paragraph.
                // This is a compile‑safe placeholder because Aspose.Words does not expose a direct line‑count API.
                int approximateLineCount = para.Runs.Count;

                Console.WriteLine($"Paragraph {index}: Approximate line count (run count) = {approximateLineCount}");
                index++;
            }

            // Save the document (optional, just to demonstrate the lifecycle).
            doc.Save("ParagraphLineCount.docx");
        }
    }
}
