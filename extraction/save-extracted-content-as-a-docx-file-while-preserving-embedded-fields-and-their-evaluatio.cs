using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with various fields.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a DATE field.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();

        // Insert a PAGE field.
        builder.InsertField(FieldType.FieldPage, true);
        builder.Writeln();

        // Insert a MERGEFIELD (will display its result after update).
        builder.InsertField(" MERGEFIELD SampleField \\* MERGEFORMAT ");
        builder.Writeln();

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedDoc = new Document(sourcePath);

        // Update all fields so their results are evaluated.
        loadedDoc.UpdateFields();

        // Clone the document to create an extracted copy.
        Document extractedDoc = (Document)loadedDoc.Clone(true);

        // Save the extracted document preserving fields and their evaluated results.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(extractedPath))
        {
            throw new InvalidOperationException("The extracted DOCX file was not created.");
        }

        // Optional: Output a simple confirmation (no interactive input required).
        Console.WriteLine("Extraction completed successfully.");
    }
}
