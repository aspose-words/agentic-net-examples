using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -------------------------
        // 1. Create two sample DOCX files with a real difference.
        // -------------------------
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Original content.");
        string doc1Path = Path.Combine(artifactsDir, "doc1.docx");
        doc1.Save(doc1Path);

        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Revised content.");
        string doc2Path = Path.Combine(artifactsDir, "doc2.docx");
        doc2.Save(doc2Path);

        // -------------------------
        // 2. Create a dummy file with an unsupported format (plain text).
        // -------------------------
        string unsupportedPath = Path.Combine(artifactsDir, "unsupported.txt");
        File.WriteAllText(unsupportedPath, "Just some text.");

        // -------------------------
        // 3. Attempt to load the unsupported file and handle the exception.
        // -------------------------
        try
        {
            // This line throws UnsupportedFileFormatException because .txt is not a Word format.
            Document unsupportedDoc = new Document(unsupportedPath);

            // If, for any reason, loading succeeded, try a comparison (won't be reached).
            doc1.Compare(unsupportedDoc, "Author", DateTime.Now);
        }
        catch (UnsupportedFileFormatException ex)
        {
            Console.WriteLine($"Caught UnsupportedFileFormatException: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions.
            Console.WriteLine($"Caught unexpected exception: {ex.Message}");
        }

        // -------------------------
        // 4. Load the supported documents and perform a normal comparison.
        // -------------------------
        Document loadedDoc1 = new Document(doc1Path);
        Document loadedDoc2 = new Document(doc2Path);

        loadedDoc1.Compare(loadedDoc2, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (loadedDoc1.Revisions.Count > 0)
        {
            Console.WriteLine($"Comparison produced {loadedDoc1.Revisions.Count} revision(s).");
        }
        else
        {
            Console.WriteLine("No revisions were produced by the comparison.");
        }

        // -------------------------
        // 5. Save the comparison result.
        // -------------------------
        string resultPath = Path.Combine(artifactsDir, "comparisonResult.docx");
        loadedDoc1.Save(resultPath);
    }
}
