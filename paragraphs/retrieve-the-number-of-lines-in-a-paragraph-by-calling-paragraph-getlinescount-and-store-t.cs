using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains manual line breaks.
        builder.Write("First line");
        builder.InsertBreak(BreakType.LineBreak);
        builder.Write("Second line");
        builder.InsertBreak(BreakType.LineBreak);
        builder.Write("Third line");

        // Get the paragraph that we have just built.
        Paragraph paragraph = builder.CurrentParagraph;

        // Approximate the number of lines in the paragraph.
        // In Aspose.Words a manual line break is represented by the Unicode character '\u000b'.
        string text = paragraph.GetText();
        int lineBreaks = 0;
        foreach (char ch in text)
        {
            if (ch == '\u000b')
                lineBreaks++;
        }
        int lineCount = lineBreaks + 1; // lines = breaks + 1

        // Store the result.
        int paragraphLineCount = lineCount;

        // Output the result (no user interaction required).
        Console.WriteLine($"Approximate line count in paragraph: {paragraphLineCount}");

        // Save the document to verify the changes.
        doc.Save("ParagraphLines.docx");
    }
}
