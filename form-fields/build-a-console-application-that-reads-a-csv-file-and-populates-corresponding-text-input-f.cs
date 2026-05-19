using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // File paths
        string templatePath = "template.docx";
        string csvPath = "data.csv";
        string outputPath = "filled.docx";

        // 1. Create a template document that contains text input form fields.
        CreateTemplate(templatePath);

        // 2. Create a sample CSV file whose header matches the form field names.
        CreateSampleCsv(csvPath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Read CSV data (first row = header, second row = values).
        List<string[]> csvRows = ReadCsv(csvPath);
        string[] headers = csvRows[0];
        string[] values = csvRows[1];

        // 5. Access the form fields collection.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validation: ensure the template contains at least one form field.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The template does not contain any form fields.");

        // 6. Populate each form field with the corresponding CSV value.
        foreach (FormField field in formFields)
        {
            // Find the column index that matches the field name (case‑insensitive).
            int columnIndex = Array.FindIndex(
                headers,
                h => string.Equals(h, field.Name, StringComparison.OrdinalIgnoreCase));

            // Validation: the CSV must contain a column for every form field.
            if (columnIndex < 0)
                throw new InvalidOperationException($"CSV does not contain a column for field '{field.Name}'.");

            // Assign the value to the text input field using the Result property.
            field.Result = values[columnIndex];
        }

        // 7. Save the populated document.
        doc.Save(outputPath);
    }

    // Creates a simple Word document with three text input form fields.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First Name:");
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();

        builder.Writeln("Last Name:");
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();

        builder.Writeln("Email:");
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();

        doc.Save(path);
    }

    // Generates a CSV file with a header that matches the form field names.
    private static void CreateSampleCsv(string path)
    {
        string[] lines =
        {
            "FirstName,LastName,Email",
            "John,Doe,john.doe@example.com"
        };
        File.WriteAllLines(path, lines);
    }

    // Reads a CSV file into a list of string arrays (simple split on commas).
    private static List<string[]> ReadCsv(string path)
    {
        var rows = new List<string[]>();
        foreach (string line in File.ReadAllLines(path))
        {
            // Note: this simple split does not handle quoted commas.
            rows.Add(line.Split(','));
        }

        if (rows.Count < 2)
            throw new InvalidOperationException("CSV must contain at least a header row and one data row.");

        return rows;
    }
}
