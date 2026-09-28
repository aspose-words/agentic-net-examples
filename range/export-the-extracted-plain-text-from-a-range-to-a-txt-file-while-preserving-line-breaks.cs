using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // Extract plain text from the whole document range.
        string extractedText = doc.Range.Text; // Preserves paragraph breaks as line breaks.

        // Define output file path.
        string outputPath = "ExtractedText.txt";

        // Write the extracted text to a .txt file.
        File.WriteAllText(outputPath, extractedText);

        // Optional: confirm that the file was created (no console output required).
        // The program ends here.
    }
}
