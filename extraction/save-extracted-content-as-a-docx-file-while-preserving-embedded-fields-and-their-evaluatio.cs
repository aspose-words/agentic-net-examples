using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various fields.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Insert a DATE field.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();

        // Insert a PAGE field.
        builder.InsertField(FieldType.FieldPage, true);
        builder.Writeln();

        // Insert a custom MERGEFIELD.
        // Use the string overload; it automatically updates the field result.
        builder.InsertField("MERGEFIELD SampleField \\* MERGEFORMAT");
        builder.Writeln();

        // Ensure all fields are evaluated.
        source.UpdateFields();

        // Save the original document.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // Load the saved document.
        Document loaded = new Document(sourcePath);

        // Extract the entire content while preserving fields and their results.
        Document extracted = loaded.Range.ToDocument();

        // Save the extracted content to a new DOCX file.
        const string extractedPath = "extracted.docx";
        extracted.Save(extractedPath);

        // Verify that the output file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
