using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Hello world!");
        builder.Writeln("This is a sample document.");
        original.Save("original.docx");

        // Create the modified document with some changes.
        Document modified = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(modified);
        builder2.Writeln("Hello world!");
        builder2.Writeln("This is an updated sample document."); // Modified line
        builder2.Writeln("Additional paragraph."); // New line
        modified.Save("modified.docx");

        // Load the original document again (or reuse the instance) and compare it with the modified one.
        // The Compare method adds revision marks to the original document.
        original.Compare(modified, "Comparer", DateTime.Now);

        // Save the document that now contains revisions highlighting the differences.
        original.Save("revisions.docx");
    }
}
