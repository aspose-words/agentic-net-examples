using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Author property (this will be displayed by the field).
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some introductory text.
        builder.Writeln("Document author:");

        // Insert an AUTHOR field that reads the Author property and update it immediately.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        authorField.Update();

        // Save the document to the local file system.
        doc.Save("AuthorField.docx");
    }
}
