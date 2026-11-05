using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first sample document.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Hello world!");
        builder1.Writeln("This is the first document.");
        doc1.Save("doc1.docx");

        // Create the second sample document with a modification.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Hello world!");
        builder2.Writeln("This is the second document with changes.");
        doc2.Save("doc2.docx");

        // Compare the two documents.
        // The comparison will generate revisions in doc1.
        string authorName = "John Doe";
        DateTime comparisonDate = new DateTime(2023, 1, 1);
        doc1.Compare(doc2, authorName, comparisonDate);

        // Save the document that now contains revision differences.
        doc1.Save("Compared.docx");

        // Output revision details to the console.
        foreach (Revision rev in doc1.Revisions)
        {
            string revisionText = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}, Text: \"{revisionText}\"");
        }
    }
}
