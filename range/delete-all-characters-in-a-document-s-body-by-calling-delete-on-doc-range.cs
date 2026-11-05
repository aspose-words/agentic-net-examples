using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This is a sample document.");

        // Save the original document locally.
        const string originalPath = "Original.docx";
        originalDoc.Save(originalPath);

        // Load the document from the saved file.
        Document doc = new Document(originalPath);

        // Delete all characters in the document's body using the Range API.
        doc.Range.Delete();

        // Save the modified document.
        const string outputPath = "DeletedContent.docx";
        doc.Save(outputPath);
    }
}
