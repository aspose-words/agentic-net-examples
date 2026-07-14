using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string unsupportedFile = Path.Combine(Directory.GetCurrentDirectory(), "unsupported.txt");
        string originalFile = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        string revisedFile = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        string resultFile = Path.Combine(Directory.GetCurrentDirectory(), "comparisonResult.docx");

        // Create a simple text file that Aspose.Words cannot load as a Word document.
        File.WriteAllText(unsupportedFile, "This is not a Word document.");

        // Attempt to load the unsupported file and handle the exception.
        try
        {
            // This line is expected to throw UnsupportedFileFormatException.
            Document unsupportedDoc = new Document(unsupportedFile);
        }
        catch (UnsupportedFileFormatException ex)
        {
            Console.WriteLine($"Caught expected exception while loading unsupported file: {ex.Message}");
        }

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        original.Save(originalFile);

        // Create the revised document with a slight change.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello Aspose.Words world.");
        revised.Save(revisedFile);

        // Load the saved documents for comparison.
        Document loadedOriginal = new Document(originalFile);
        Document loadedRevised = new Document(revisedFile);

        // Perform the comparison. Revisions will be added to the original document.
        loadedOriginal.Compare(loadedRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (loadedOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result.
        loadedOriginal.Save(resultFile);

        Console.WriteLine($"Comparison completed. Revisions count: {loadedOriginal.Revisions.Count}");
        Console.WriteLine($"Result saved to: {resultFile}");
    }
}
