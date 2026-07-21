using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with a few fields.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a DATE field and a PAGE field.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();
        builder.InsertField(FieldType.FieldPage, true);
        builder.Writeln();

        // Save the source document to disk.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document back from the file system.
        Document loadedDoc = new Document(sourcePath);

        // Update all fields so that their results are evaluated.
        loadedDoc.UpdateFields();

        // Extract the entire content (including fields and their results) into a new document.
        Document extractedDoc = loadedDoc.Range.ToDocument();

        // Save the extracted content as a new DOCX file.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Verify that the output file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");

        // Optional: output a confirmation message.
        Console.WriteLine("Extraction completed successfully. Extracted file: " + extractedPath);
    }
}
