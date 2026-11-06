using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(original);
        originalBuilder.Writeln("This is the original paragraph.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder revisedBuilder = new DocumentBuilder(revised);
        revisedBuilder.Writeln("This is the revised paragraph with extra text.");

        // Define a custom author name and a specific timestamp.
        string customAuthor = "John Doe";
        DateTime customTimestamp = new DateTime(2023, 1, 1, 12, 0, 0, DateTimeKind.Utc);

        // Compare the documents, attributing revisions to the custom author and timestamp.
        original.Compare(revised, customAuthor, customTimestamp);

        // Output revision details to the console.
        foreach (Revision revision in original.Revisions)
        {
            Console.WriteLine($"Revision Type: {revision.RevisionType}, Author: {revision.Author}, Date: {revision.DateTime:u}");
        }

        // Save the resulting document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(outputPath);
    }
}
