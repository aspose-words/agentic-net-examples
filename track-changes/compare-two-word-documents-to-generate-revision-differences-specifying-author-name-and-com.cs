using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = "Original.docx";
        string editedPath = "Edited.docx";
        string resultPath = "Compared.docx";

        // -----------------------------------------------------------------
        // Create the original document.
        // -----------------------------------------------------------------
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the original document.");
        builder.Writeln("It contains a few lines of text.");
        originalDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // Create the edited document with some changes.
        // -----------------------------------------------------------------
        Document editedDoc = new Document();
        builder = new DocumentBuilder(editedDoc);
        builder.Writeln("This is the edited document."); // changed line
        builder.Writeln("It contains a few lines of text."); // unchanged line
        builder.Writeln("An additional line was added."); // new line
        editedDoc.Save(editedPath);

        // -----------------------------------------------------------------
        // Load the original document (could also reuse the instance).
        // -----------------------------------------------------------------
        Document compareDoc = new Document(originalPath);

        // Perform the comparison. The revisions will be added to compareDoc.
        string authorName = "Comparer";
        DateTime compareDate = DateTime.Now;
        compareDoc.Compare(editedDoc, authorName, compareDate);

        // -----------------------------------------------------------------
        // Output revision details to the console.
        // -----------------------------------------------------------------
        Console.WriteLine("Revisions found after comparison:");
        foreach (Revision rev in compareDoc.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}");
            Console.WriteLine($"  Author: {rev.Author}");
            Console.WriteLine($"  Date: {rev.DateTime}");
            Console.WriteLine($"  Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // -----------------------------------------------------------------
        // Save the document that now contains the revision markup.
        // -----------------------------------------------------------------
        compareDoc.Save(resultPath);
    }
}
