using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Author property – the AUTHOR field will read this value.
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Initialize a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that introduces the author field.
        builder.Writeln("Document author:");

        // Insert an AUTHOR field and update it immediately so it shows the current author.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        authorField.Update();

        // Save the document to the local file system.
        doc.Save("AuthorField.docx");
    }
}
