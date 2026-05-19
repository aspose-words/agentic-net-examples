using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs to generate line breaks.
        builder.Writeln("First line.");
        builder.Writeln("Second line.");
        builder.Writeln("Third line.");

        // Extract the plain text from the whole‑document range.
        string extractedText = doc.Range.Text;

        // Path for the output .txt file.
        string outputFile = "ExtractedText.txt";

        // Write the extracted text to the file, preserving the original line breaks.
        File.WriteAllText(outputFile, extractedText);

        // Simple verification that the file was created.
        if (File.Exists(outputFile))
        {
            Console.WriteLine($"Extracted text saved to '{outputFile}'.");
        }
    }
}
