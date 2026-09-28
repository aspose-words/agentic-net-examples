using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class DocumentComparisonExample
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparison-result.docx");

        // Create the original document with some content.
        Document originalDoc = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(originalDoc);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a few lines of text.");
        originalDoc.Save(originalPath, SaveFormat.Docx);

        // Create the revised document with differences.
        Document revisedDoc = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revisedDoc);
        builderRevised.Writeln("This is the revised document."); // Changed line.
        builderRevised.Writeln("It contains a few lines of text."); // Same line.
        builderRevised.Writeln("An additional line was added."); // New line.
        revisedDoc.Save(revisedPath, SaveFormat.Docx);

        // Load the documents from disk.
        Document loadedOriginal = new Document(originalPath);
        Document loadedRevised = new Document(revisedPath);

        // Perform comparison. Revisions will be added to loadedOriginal.
        loadedOriginal.Compare(loadedRevised, "ComparisonAuthor", DateTime.Now);

        // Verify that revisions were created.
        if (loadedOriginal.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result (includes revisions).
        loadedOriginal.Save(resultPath, SaveFormat.Docx);
    }
}
