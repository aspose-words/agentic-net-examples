using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some paragraphs – each WriteLine inserts a paragraph break.
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // Extract the plain text of the whole document range.
        // The returned string contains control characters such as '\r' for paragraph breaks.
        string extractedText = doc.Range.Text;

        // Define the output file path (in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedText.txt");

        // Write the extracted text to a .txt file, preserving the line breaks.
        // Using UTF8 encoding ensures all characters are saved correctly.
        File.WriteAllText(outputPath, extractedText, System.Text.Encoding.UTF8);

        // Optionally, demonstrate saving via TxtSaveOptions (produces the same result).
        // TxtSaveOptions options = new TxtSaveOptions();
        // doc.Save(outputPath, options);
    }
}
