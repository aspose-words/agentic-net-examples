using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs of varying length.
        builder.Writeln("Short paragraph.");
        builder.Writeln("This is a longer paragraph that will likely wrap onto multiple lines when rendered in a typical Word document layout. It contains several sentences to increase its length.");
        builder.Writeln("Another paragraph.\nIt even contains an explicit line break within the text, which should be treated as a separate line in the layout.");
        builder.Writeln("Final paragraph with enough content to span several lines. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Save the document (optional, just to demonstrate the lifecycle).
        doc.Save("ParagraphLineCount.docx");

        // Approximate line count per paragraph.
        // For demonstration we assume an average of 80 characters per line.
        const int avgCharsPerLine = 80;

        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Get the paragraph text without the terminating paragraph break.
            string text = para.GetText().TrimEnd('\r', '\n');

            // Simple approximation: number of lines = ceil(text length / avgCharsPerLine)
            int approxLineCount = (int)Math.Ceiling((double)text.Length / avgCharsPerLine);
            // Ensure at least one line for empty paragraphs.
            if (approxLineCount == 0) approxLineCount = 1;

            Console.WriteLine($"Paragraph {para.GetText().TrimEnd('\r', '\n').Substring(0, Math.Min(30, text.Length))}... : Approx. {approxLineCount} line(s)");
        }
    }
}
