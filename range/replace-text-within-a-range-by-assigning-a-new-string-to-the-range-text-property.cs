using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the original text.");

        // Save the original document (optional, just to demonstrate loading).
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(originalPath);

        // Remove all existing content from the document.
        loadedDoc.Range.Delete();

        // Insert the replacement text into the now‑empty document.
        DocumentBuilder replaceBuilder = new DocumentBuilder(loadedDoc);
        replaceBuilder.Writeln("This is the replaced text.");

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        loadedDoc.Save(modifiedPath);
    }
}
