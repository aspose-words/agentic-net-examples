using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the built‑in Author property (this will be displayed by the field).
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Add a paragraph introducing the author field.
        builder.Writeln("Document Author:");

        // Insert a DOCPROPERTY field that dynamically shows the Author property.
        builder.InsertField("DOCPROPERTY Author \\* MERGEFORMAT", string.Empty);

        // Save the document to disk.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
