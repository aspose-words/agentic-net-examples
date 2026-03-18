using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that introduces the author field.
        builder.Writeln("This document was authored by:");

        // Insert an AUTHOR field. The second argument (true) updates the field immediately.
        // The field reads the value of the built‑in Author property.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);

        // If the Author property is empty, provide a fallback value.
        if (string.IsNullOrEmpty(doc.BuiltInDocumentProperties.Author))
        {
            doc.FieldOptions.DefaultDocumentAuthor = "Default Author";
            authorField.Update(); // Refresh the field to display the fallback.
        }

        // Save the document to disk.
        doc.Save("AuthorField.docx");
    }
}
