using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths
        string originalPath = Path.Combine(Environment.CurrentDirectory, "original.docx");
        string resultPath = Path.Combine(Environment.CurrentDirectory, "result.docx");

        // -------------------------------------------------
        // Create a new document with a bookmark named "Draft"
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some text before the bookmark
        builder.Writeln("This is the beginning of the document.");

        // Start the bookmark, insert content, then end the bookmark
        builder.StartBookmark("Draft");
        builder.Writeln("This content is inside the Draft bookmark and will be removed.");
        builder.EndBookmark("Draft");

        // Insert some text after the bookmark
        builder.Writeln("This is the end of the document.");

        // Save the original document
        doc.Save(originalPath);

        // -------------------------------------------------
        // Load the document and remove the "Draft" bookmark
        // -------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Remove the bookmark named "Draft" (and its content as per task requirement)
        loadedDoc.Range.Bookmarks.Remove("Draft");

        // Save the modified document
        loadedDoc.Save(resultPath);
    }
}
