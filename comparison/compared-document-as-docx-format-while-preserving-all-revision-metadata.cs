using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Create temporary file paths.
        string originalPath = Path.Combine(Path.GetTempPath(), "Original.docx");
        string editedPath   = Path.Combine(Path.GetTempPath(), "Edited.docx");
        string resultPath   = Path.Combine(Path.GetTempPath(), "ComparedResult.docx");

        // Build the original document.
        var originalDoc = new Document();
        var builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world!");
        originalDoc.Save(originalPath);

        // Build the edited document (adds an extra line).
        var editedDoc = new Document();
        var editedBuilder = new DocumentBuilder(editedDoc);
        editedBuilder.Writeln("Hello world!");
        editedBuilder.Writeln("Additional line.");
        editedDoc.Save(editedPath);

        // Load the documents.
        var docOriginal = new Document(originalPath);
        var docEdited   = new Document(editedPath);

        // Compare edited document against the original; revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now);

        // Save the comparison result while preserving all revision metadata.
        docOriginal.Save(resultPath);

        Console.WriteLine($"Comparison completed. Result saved to: {resultPath}");
    }
}
