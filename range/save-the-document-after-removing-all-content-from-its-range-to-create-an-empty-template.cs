using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is some sample text that will be removed.");

        // Verify that the document currently contains text.
        // (Optional, just for demonstration; can be omitted in production.)
        Console.WriteLine("Before deletion: " + doc.GetText().Trim());

        // Delete all characters in the whole document range, leaving an empty template.
        doc.Range.Delete();

        // Verify that the document is now empty.
        Console.WriteLine("After deletion: '" + doc.GetText().Trim() + "'");

        // Define the output file path (in the same directory as the executable).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmptyTemplate.docx");

        // Save the empty document.
        doc.Save(outputPath);

        // Inform the user where the file was saved.
        Console.WriteLine("Empty template saved to: " + outputPath);
    }
}
