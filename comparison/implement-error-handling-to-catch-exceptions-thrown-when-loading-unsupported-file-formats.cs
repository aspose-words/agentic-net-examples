using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonOutput");
        Directory.CreateDirectory(outputDir);

        // Create first sample document
        string originalPath = Path.Combine(outputDir, "Original.docx");
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("Hello world.");
        originalDoc.Save(originalPath);

        // Create second sample document with a difference
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("Hello revised world.");
        revisedDoc.Save(revisedPath);

        // Load the documents for comparison
        Document original = new Document(originalPath);
        Document revised = new Document(revisedPath);

        // Perform comparison
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save comparison result
        string resultPath = Path.Combine(outputDir, "Compared.docx");
        original.Save(resultPath);
        Console.WriteLine($"Comparison completed. Revisions count: {original.Revisions.Count}");

        // Attempt to load an unsupported file format
        string txtPath = Path.Combine(outputDir, "Unsupported.txt");
        File.WriteAllText(txtPath, "Just some plain text.");

        try
        {
            Document unsupported = new Document(txtPath);
            // If no exception, indicate unexpected success
            Console.WriteLine("Unexpectedly loaded unsupported file without exception.");
        }
        catch (Exception ex)
        {
            // Aspose.Words throws a generic exception for unsupported formats in some versions.
            // We treat any exception here as an indication that the format is not supported.
            Console.WriteLine($"Caught exception while loading unsupported file: {ex.Message}");
        }
    }
}
