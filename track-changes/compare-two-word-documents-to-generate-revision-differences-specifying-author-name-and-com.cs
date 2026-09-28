using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world!");
        builder.Writeln("This is the original document.");
        originalDoc.Save("Original.docx");

        // Create the modified document.
        Document modifiedDoc = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(modifiedDoc);
        builder2.Writeln("Hello world!");
        builder2.Writeln("This is the modified document with extra line.");
        builder2.Writeln("Additional paragraph added.");
        modifiedDoc.Save("Modified.docx");

        // Compare the documents to generate revisions, specifying author and date.
        string authorName = "ComparisonAuthor";
        DateTime comparisonDate = new DateTime(2023, 1, 1, 12, 0, 0);
        originalDoc.Compare(modifiedDoc, authorName, comparisonDate);

        // Save the document that now contains revision differences.
        originalDoc.Save("ComparedWithRevisions.docx");

        // Output revision details.
        foreach (Revision rev in originalDoc.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
