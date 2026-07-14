using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Properties;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Author property (this value will be shown by the field).
        doc.BuiltInDocumentProperties.Author = "Alice Example";

        // Provide a fallback author name in case the Author property is empty.
        doc.FieldOptions.DefaultDocumentAuthor = "Default Author";

        // Use DocumentBuilder to add a paragraph and the AUTHOR field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document created by:");

        // Insert an AUTHOR field that displays the current Author property.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        authorField.Update(); // Ensure the field result is up‑to‑date.

        // Save the document.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);
        string outputPath = System.IO.Path.Combine(outputDir, "AuthorField.docx");
        doc.Save(outputPath);
    }
}
