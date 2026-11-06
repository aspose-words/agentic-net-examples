using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first (original) document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("This paragraph will stay the same.");
        builder.Writeln("This is the third paragraph.");
        originalDoc.Save("Original.docx");

        // Create the second (modified) document.
        Document modifiedDoc = new Document();
        builder = new DocumentBuilder(modifiedDoc);
        builder.Writeln("This is the first paragraph."); // unchanged
        builder.Writeln("This paragraph has been modified."); // changed text
        builder.Writeln("This is the third paragraph."); // unchanged
        builder.Writeln("An additional paragraph is added."); // new paragraph
        modifiedDoc.Save("Modified.docx");

        // Load the documents for comparison.
        Document docToCompare = new Document("Original.docx");
        Document docModified = new Document("Modified.docx");

        // Compare the documents. The result (revisions) are stored in docToCompare.
        docToCompare.Compare(docModified, "Comparer", DateTime.Now);

        // Save the comparison result which contains revision marks.
        docToCompare.Save("ComparisonResult.docx");

        // Output revision information to the console.
        RevisionCollection revisions = docToCompare.Revisions;
        Console.WriteLine($"Total revisions detected: {revisions.Count}");
        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
