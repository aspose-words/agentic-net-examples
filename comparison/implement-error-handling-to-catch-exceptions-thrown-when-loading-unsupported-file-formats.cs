using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create two sample documents with a real difference ----------
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("This is the original document.");

        Document editedDoc = new Document();
        DocumentBuilder builderEdited = new DocumentBuilder(editedDoc);
        builderEdited.Writeln("This is the edited document with a change.");

        // ---------- Perform comparison ----------
        // The original document will receive revision marks after comparison.
        originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);
        string comparisonResultPath = Path.Combine(artifactsDir, "ComparisonResult.docx");
        originalDoc.Save(comparisonResultPath);

        // ---------- Attempt to load an unsupported file format ----------
        // Create a dummy file with an unknown extension.
        string unsupportedFilePath = Path.Combine(artifactsDir, "unsupported.xyz");
        File.WriteAllText(unsupportedFilePath, "This content is not in a supported Word format.");

        try
        {
            // This line is expected to throw UnsupportedFileFormatException.
            Document unsupportedDoc = new Document(unsupportedFilePath);
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Handle the specific exception for unsupported formats.
            Console.WriteLine($"Unsupported file format encountered: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Fallback for any other unexpected exceptions.
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
