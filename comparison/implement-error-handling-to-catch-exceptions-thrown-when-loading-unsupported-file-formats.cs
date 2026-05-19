using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Define a working directory for all artifacts.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the sample documents.
        string originalPath = Path.Combine(workDir, "original.docx");
        string revisedPath = Path.Combine(workDir, "revised.docx");
        string resultPath = Path.Combine(workDir, "comparison_result.docx");
        string unsupportedPath = Path.Combine(workDir, "unsupported.xyz");

        // -----------------------------------------------------------------
        // 1. Create two simple documents with a deterministic difference.
        // -----------------------------------------------------------------
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world! This is the original document.");

        original.Save(originalPath);

        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world! This is the revised document with a change.");

        revised.Save(revisedPath);

        // -----------------------------------------------------------------
        // 2. Attempt to load an unsupported file format and handle the exception.
        // -----------------------------------------------------------------
        try
        {
            // This file does not exist and has an unknown extension, triggering an exception.
            // The Document constructor will attempt to detect the format and fail.
            Document unsupported = new Document(unsupportedPath);
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Log the exception details – in a real scenario you might take corrective action.
            Console.WriteLine($"Unsupported file format encountered: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions.
            Console.WriteLine($"Unexpected error while loading document: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 3. Load the previously created valid documents.
        // -----------------------------------------------------------------
        Document docOriginal = new Document(originalPath);
        Document docRevised = new Document(revisedPath);

        // Ensure there are no pre-existing revisions before comparison.
        if (docOriginal.HasRevisions || docRevised.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // -----------------------------------------------------------------
        // 4. Perform the comparison, producing revisions in the original document.
        // -----------------------------------------------------------------
        docOriginal.Compare(docRevised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (docOriginal.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // -----------------------------------------------------------------
        // 5. Save the comparison result.
        // -----------------------------------------------------------------
        docOriginal.Save(resultPath);

        // Output a simple summary to the console.
        Console.WriteLine($"Comparison completed. Revisions found: {docOriginal.Revisions.Count}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
