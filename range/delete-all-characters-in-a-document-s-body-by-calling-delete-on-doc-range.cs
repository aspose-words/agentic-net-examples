using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text to the document body.
        builder.Writeln("Hello World!");
        builder.Writeln("This is a second line.");

        // Save the original document for reference.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // Delete all characters in the whole document by using the Range.Delete method.
        doc.Range.Delete();

        // Save the document after the deletion.
        string deletedPath = Path.Combine(outputDir, "Deleted.docx");
        doc.Save(deletedPath);

        // Optional: write a simple verification to the console.
        Console.WriteLine("Original document saved to: " + originalPath);
        Console.WriteLine("Document after Range.Delete saved to: " + deletedPath);
        Console.WriteLine("Document text after deletion (should be empty):");
        Console.WriteLine($"\"{doc.GetText().Trim()}\"");
    }
}
