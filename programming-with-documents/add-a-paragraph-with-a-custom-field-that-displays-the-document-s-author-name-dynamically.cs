using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the built‑in Author property of the document.
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph.
        builder.Writeln();

        // Insert a field that displays the Author property dynamically.
        // The field code "DOCPROPERTY Author" pulls the current Author value.
        builder.InsertField(" DOCPROPERTY Author ");

        // Add a line break after the field (optional).
        builder.Writeln();

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
