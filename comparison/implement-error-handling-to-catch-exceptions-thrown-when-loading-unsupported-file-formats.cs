using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("Original content.");
        string originalPath = Path.Combine(artifactsDir, "original.docx");
        originalDoc.Save(originalPath);

        // Create the revised document with a deliberate difference.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("Revised content.");
        string revisedPath = Path.Combine(artifactsDir, "revised.docx");
        revisedDoc.Save(revisedPath);

        // Create a file that Aspose.Words cannot recognise as a supported format.
        string unsupportedPath = Path.Combine(artifactsDir, "unsupported.bin");
        File.WriteAllBytes(unsupportedPath, new byte[] { 0x00, 0x01, 0x02, 0x03 });

        // Attempt to load the unsupported file and handle the expected exception.
        try
        {
            // This constructor call should throw UnsupportedFileFormatException.
            Document unsupportedDoc = new Document(unsupportedPath);
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Log the exception message – the program continues after handling.
            Console.WriteLine($"Caught unsupported format exception: {ex.Message}");
        }

        // Load the valid documents for comparison.
        Document original = new Document(originalPath);
        Document revised = new Document(revisedPath);

        // Perform the comparison; revisions will be added to the original document.
        original.Compare(revised, "DemoAuthor", DateTime.Now);

        // Verify that at least one revision was produced.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison.");
        }

        // Save the comparison result.
        string resultPath = Path.Combine(artifactsDir, "comparisonResult.docx");
        original.Save(resultPath);
    }
}
