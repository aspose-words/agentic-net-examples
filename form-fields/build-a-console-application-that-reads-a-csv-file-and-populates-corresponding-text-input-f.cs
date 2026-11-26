using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file paths
        string csvPath = "data.csv";
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // Create a sample CSV file
        File.WriteAllText(csvPath, "Name,Email\nJohn Doe,john.doe@example.com");

        // Create a template document with text input form fields
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a text input field for Name
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "Default Name", 0);
        builder.Writeln();

        // Insert a text input field for Email
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "default@example.com", 0);
        builder.Writeln();

        // Save the template document
        templateDoc.Save(templatePath);

        // Load the template document for processing
        Document doc = new Document(templatePath);

        // Read CSV content
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length < 2)
            throw new InvalidOperationException("CSV file must contain header and at least one data row.");

        // Parse header and first data row
        string[] headers = csvLines[0].Split(',');
        string[] values = csvLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header and data column counts do not match.");

        // Populate form fields with CSV values
        for (int i = 0; i < headers.Length; i++)
        {
            string fieldName = headers[i].Trim();
            string fieldValue = values[i].Trim();

            FormField field = doc.Range.FormFields[fieldName];
            if (field == null)
                throw new InvalidOperationException($"Form field '{fieldName}' not found in the document.");

            // Only text input fields have Result property
            field.Result = fieldValue;

            // Validate assignment
            if (field.Result != fieldValue)
                throw new InvalidOperationException($"Failed to set value for field '{fieldName}'.");
        }

        // Save the populated document
        doc.Save(outputPath);
    }
}
