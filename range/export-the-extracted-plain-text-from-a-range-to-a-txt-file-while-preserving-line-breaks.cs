using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // Extract plain text from the whole document range.
        string extractedText = doc.Range.Text;

        // Save the extracted text to a .txt file, preserving line breaks.
        string outputPath = "ExtractedText.txt";
        File.WriteAllText(outputPath, extractedText);

        // Optional: confirm the file was created (no console output required).
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the output text file.");
        }
    }
}
