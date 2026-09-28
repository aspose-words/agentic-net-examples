using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the template, CSV data, and output document.
        const string templatePath = "Template.docx";
        const string csvPath = "Data.csv";
        const string outputPath = "Output.docx";

        // Step 1: Create a template document with text input form fields.
        CreateTemplateDocument(templatePath);

        // Step 2: Create a sample CSV file that matches the form field names.
        CreateSampleCsv(csvPath);

        // Step 3: Load the template document.
        Document doc = new Document(templatePath);

        // Step 4: Read CSV data.
        Dictionary<string, string> fieldValues = ReadCsvIntoDictionary(csvPath);

        // Step 5: Populate form fields with CSV values.
        foreach (KeyValuePair<string, string> kvp in fieldValues)
        {
            // Validate that the form field exists.
            FormField? formField = doc.Range.FormFields[kvp.Key];
            if (formField == null)
                throw new InvalidOperationException($"Form field '{kvp.Key}' not found in the document.");

            // Assign the CSV value to the text input field.
            formField.Result = kvp.Value ?? string.Empty;

            // Optional validation: ensure the value was set.
            if (formField.Result != kvp.Value)
                throw new InvalidOperationException($"Failed to set value for field '{kvp.Key}'.");
        }

        // Step 6: Save the populated document.
        doc.Save(outputPath);
    }

    // Creates a DOCX file with three text input form fields: FirstName, LastName, Email.
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input field for FirstName.
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "John", 0);
        builder.Writeln();

        // Insert a text input field for LastName.
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "Doe", 0);
        builder.Writeln();

        // Insert a text input field for Email.
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "example@example.com", 0);
        builder.Writeln();

        doc.Save(path);
    }

    // Writes a simple CSV file with a header matching the form field names and one data row.
    private static void CreateSampleCsv(string path)
    {
        string[] lines =
        {
            "FirstName,LastName,Email",
            "Jane,Smith,jane.smith@example.com"
        };
        File.WriteAllLines(path, lines);
    }

    // Reads the CSV file and returns a dictionary mapping column names to their values.
    private static Dictionary<string, string> ReadCsvIntoDictionary(string path)
    {
        string[] allLines = File.ReadAllLines(path);
        if (allLines.Length < 2)
            throw new InvalidOperationException("CSV file must contain at least a header and one data row.");

        // Parse header.
        string[] headers = allLines[0].Split(',');

        // Parse first data row.
        string[] values = allLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header and data column counts do not match.");

        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++)
        {
            string key = headers[i].Trim();
            string value = values[i].Trim();
            dict[key] = value;
        }

        return dict;
    }
}
