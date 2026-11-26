using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");

        // -----------------------------------------------------------------
        // Create the original document.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello world!");
        builder.Writeln("This is the original document.");
        originalDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the revised document with some changes.
        // -----------------------------------------------------------------
        Document revisedDoc = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revisedDoc);
        builder2.Writeln("Hello world!"); // unchanged line
        builder2.Writeln("This is the modified document."); // changed line
        builder2.Writeln("Additional line."); // new line
        revisedDoc.Save(revisedPath);

        // -----------------------------------------------------------------
        // Load the original document again (could reuse originalDoc) and compare.
        // The Compare method adds revisions to the original document.
        // -----------------------------------------------------------------
        Document compareDoc = new Document(originalPath);
        compareDoc.Compare(revisedDoc, "Comparer", DateTime.Now);
        compareDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // Inspect the revisions created by the comparison.
        // -----------------------------------------------------------------
        RevisionCollection revisions = compareDoc.Revisions;
        Console.WriteLine($"Total revisions detected: {revisions.Count}");

        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // The program finishes here without waiting for user input.
    }
}
