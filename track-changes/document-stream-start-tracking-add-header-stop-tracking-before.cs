using System;
using System.IO;
using Aspose.Words;

public class RevisionExample
{
    /// <summary>
    /// Loads a Word document from a stream, starts revision tracking, adds a header,
    /// stops tracking, and saves the result to the specified file.
    /// </summary>
    /// <param name="inputStream">Stream containing the source document.</param>
    /// <param name="outputPath">Full path where the modified document will be saved.</param>
    public static void ProcessDocument(Stream inputStream, string outputPath)
    {
        // Load the document from the provided stream (Document(Stream) constructor is the approved load rule).
        Document doc = new Document(inputStream);

        // Begin tracking revisions. All subsequent changes will be recorded as revisions.
        doc.StartTrackRevisions("Author");

        // Use DocumentBuilder to insert a header into the first section.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("My Header");

        // Stop tracking revisions so further edits are not recorded.
        doc.StopTrackRevisions();

        // Save the modified document to the given file path (Document.Save(string) is the approved save rule).
        doc.Save(outputPath);
    }

    // ---------------------------------------------------------------------
    // Entry point required for a console application.
    // ---------------------------------------------------------------------
    public static void Main(string[] args)
    {
        // Simple argument handling – expects input file path and output file path.
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: RevisionExample <input-doc-path> <output-doc-path>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Open the source document as a read‑only stream.
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            ProcessDocument(stream, outputPath);
        }

        Console.WriteLine($"Document processed and saved to '{outputPath}'.");
    }
}
