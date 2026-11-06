using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the two documents.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the comparison result.
        string comparePath = Path.Combine(Directory.GetCurrentDirectory(), "comparison.docx");
        original.Save(comparePath);

        // Create a plain text file that Aspose.Words cannot load.
        string unsupportedPath = Path.Combine(Directory.GetCurrentDirectory(), "unsupported.txt");
        File.WriteAllText(unsupportedPath, "This is not a supported document format.");

        // Attempt to load the unsupported file and handle any exception.
        try
        {
            Document unsupportedDoc = new Document(unsupportedPath);
            // If loading somehow succeeds, attempt a comparison (for demonstration).
            original.Compare(unsupportedDoc, "Comparer", DateTime.Now);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Caught exception while loading unsupported file: {ex.Message}");
        }
    }
}
