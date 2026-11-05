using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the author property (this can be changed later and the field will update).
        doc.BuiltInDocumentProperties.Author = "John Doe";

        // Add a paragraph and insert an AUTHOR field that displays the document's author dynamically.
        builder.Writeln("Document Author:");
        FieldAuthor authorField = (FieldAuthor)builder.InsertField(FieldType.FieldAuthor, true);
        authorField.Update();

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
