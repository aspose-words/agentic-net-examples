using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Paths for the sample documents.
        const string originalPath = "Original.docx";
        const string revisedPath = "Revised.docx";
        const string comparedPath = "Compared.docx";

        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This is the second paragraph.");
        originalDoc.Save(originalPath);

        // Create the revised document with some changes.
        Document revisedDoc = new Document();
        builder = new DocumentBuilder(revisedDoc);
        builder.Writeln("This is the first paragraph with an edit."); // Modified text.
        builder.Writeln("This is the second paragraph."); // Same as original.
        builder.Writeln("This is an added third paragraph."); // New paragraph.
        revisedDoc.Save(revisedPath);

        // Load the original document again to perform comparison.
        Document compareDoc = new Document(originalPath);

        // Compare the revised document against the original.
        // The comparison creates revision marks in compareDoc.
        string authorName = "John Doe";
        DateTime comparisonDate = DateTime.Now;
        compareDoc.Compare(revisedDoc, authorName, comparisonDate);
        compareDoc.Save(comparedPath);

        // Output revision details.
        Console.WriteLine("Revisions found after comparison:");
        foreach (Revision rev in compareDoc.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}");
            Console.WriteLine($"  Author: {rev.Author}");
            Console.WriteLine($"  Date: {rev.DateTime}");
            Console.WriteLine($"  Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Optionally accept all revisions to produce a clean document.
        // compareDoc.AcceptAllRevisions();
        // compareDoc.Save("Accepted.docx");
    }
}
