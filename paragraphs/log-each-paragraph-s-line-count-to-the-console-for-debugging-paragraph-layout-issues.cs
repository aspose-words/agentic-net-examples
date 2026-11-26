using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ParagraphLineCountDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add sample paragraphs with varying line breaks.
            builder.Writeln("First line.\nSecond line.");          // 2 lines
            builder.Writeln("Single line paragraph.");            // 1 line
            builder.Writeln("Line1.\nLine2.\nLine3.");            // 3 lines
            builder.Writeln("Line A.\r\nLine B.\r\nLine C.");    // 3 lines using CRLF

            // Save the document (optional for visual verification).
            doc.Save("Output.docx");

            // Traverse all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            int index = 1;
            foreach (Paragraph paragraph in paragraphs)
            {
                // Approximate line count by counting line‑break characters in the paragraph text.
                // Paragraph.GetText() returns the text including paragraph mark characters.
                string text = paragraph.GetText();

                // Split on both '\r' and '\n' and remove empty entries.
                int lineCount = string.IsNullOrEmpty(text)
                    ? 0
                    : text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

                Console.WriteLine($"Paragraph {index}: Approximate line count = {lineCount}");
                index++;
            }
        }
    }
}
