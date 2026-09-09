using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Author property (optional, otherwise the default author will be used).
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that introduces the author field.
        builder.Writeln("Document author:");

        // Insert an AUTHOR field that displays the document's author dynamically.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        // Update the field to reflect the current Author property.
        authorField.Update();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AuthorField.docx");
        doc.Save(outputPath);
    }
}
