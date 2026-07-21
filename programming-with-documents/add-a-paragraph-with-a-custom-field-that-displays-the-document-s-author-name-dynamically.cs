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

        // Set the built‑in Author property (this value will be shown by the field).
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Initialize a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Write("Document author: ");

        // Insert an AUTHOR field and update it so it displays the current author.
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        authorField.Update();

        // Finish the paragraph.
        builder.Writeln();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "AuthorField.docx");
        doc.Save(outputPath);
    }
}
