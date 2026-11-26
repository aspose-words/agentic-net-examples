using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Use DocumentBuilder to add some paragraphs with line breaks.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // Extract plain text from the whole document range.
        string plainText = doc.Range.Text;

        // Save the extracted text to a .txt file, preserving line breaks.
        string outputPath = "ExtractedText.txt";
        File.WriteAllText(outputPath, plainText);

        // Optional: confirm that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Text successfully exported to '{outputPath}'.");
        }
    }
}
