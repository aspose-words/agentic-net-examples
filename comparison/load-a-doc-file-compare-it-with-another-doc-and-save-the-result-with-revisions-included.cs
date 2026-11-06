using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document originalDoc = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(originalDoc);
        originalBuilder.Writeln("This is the original document.");
        originalBuilder.Writeln("It contains a line that will be changed.");

        // Save the original document to a local file.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        originalDoc.Save(originalPath, SaveFormat.Docx);

        // Create the revised document with a modification.
        Document revisedDoc = new Document();
        DocumentBuilder revisedBuilder = new DocumentBuilder(revisedDoc);
        revisedBuilder.Writeln("This is the original document.");
        revisedBuilder.Writeln("It contains a line that has been changed."); // Modified line.

        // Save the revised document to a local file.
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        revisedDoc.Save(revisedPath, SaveFormat.Docx);

        // Load the documents back (demonstrates loading).
        Document original = new Document(originalPath);
        Document revised = new Document(revisedPath);

        // Compare the documents, providing author name and comparison date.
        original.Compare(revised, "ComparisonAuthor", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after comparison.");
        }

        // Save the comparison result which includes the revisions.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(resultPath, SaveFormat.Docx);
    }
}
