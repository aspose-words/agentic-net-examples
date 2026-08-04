using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs and a page break to generate line breaks in the text.
        builder.Writeln("First line.");
        builder.Writeln("Second line.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third line after page break.");

        // Extract the plain text from the whole‑document range.
        // The Range.Text property includes control characters such as carriage returns and page breaks.
        string extractedText = doc.Range.Text;

        // Save the extracted text to a .txt file, preserving the line breaks as they appear in the string.
        string outputPath = "ExtractedText.txt";
        File.WriteAllText(outputPath, extractedText);
    }
}
