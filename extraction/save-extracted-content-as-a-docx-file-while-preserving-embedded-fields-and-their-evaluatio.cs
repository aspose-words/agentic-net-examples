using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with fields.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a DATE field.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();

        // Insert a PAGE field.
        builder.InsertField(FieldType.FieldPage, true);
        builder.Writeln();

        // Insert a QUOTE field that displays static text.
        builder.InsertField(" QUOTE  \"Extracted content\" ");
        builder.Writeln();

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document from the file.
        Document loadedDoc = new Document(sourcePath);

        // Update all fields so their results are current.
        loadedDoc.UpdateFields();

        // Extract the entire content preserving fields and their results.
        Document extractedDoc = loadedDoc.Range.ToDocument();

        // Save the extracted document.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Validate that the output file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");

        // Confirmation message.
        Console.WriteLine("Extraction completed successfully. File saved as " + extractedPath);
    }
}
