using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a slight change.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Save both documents to the local file system.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        original.Save(originalPath);
        revised.Save(revisedPath);

        // Create a dummy file with an unsupported format (plain text).
        string unsupportedPath = Path.Combine(Directory.GetCurrentDirectory(), "unsupported.txt");
        File.WriteAllText(unsupportedPath, "Just some plain text.");

        // Attempt to load the unsupported file as a Word document.
        try
        {
            // This line is expected to throw UnsupportedFileFormatException.
            Document unsupportedDoc = new Document(unsupportedPath);
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Handle the exception gracefully and report it.
            Console.WriteLine($"Caught UnsupportedFileFormatException: {ex.Message}");
        }

        // Load the previously saved valid documents.
        Document loadedOriginal = new Document(originalPath);
        Document loadedRevised = new Document(revisedPath);

        // Perform the comparison. Revisions will be added to the original document.
        loadedOriginal.Compare(loadedRevised, "Author", DateTime.Now);

        // Verify that revisions were created.
        if (loadedOriginal.Revisions.Count > 0)
        {
            Console.WriteLine($"Comparison produced {loadedOriginal.Revisions.Count} revision(s).");
        }
        else
        {
            Console.WriteLine("No revisions were detected after comparison.");
        }

        // Save the comparison result to a new file.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparisonResult.docx");
        loadedOriginal.Save(resultPath);
    }
}
