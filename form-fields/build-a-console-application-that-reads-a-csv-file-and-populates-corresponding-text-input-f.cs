using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        const string templatePath = "Template.docx";
        const string csvPath = "Data.csv";
        const string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template with text input form fields if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Please fill the form below:");
            // Insert a text input field for Name.
            builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 50);
            builder.Writeln(); // Move to next line.
            // Insert a text input field for Email.
            builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 50);
            builder.Writeln();
            // Insert a text input field for Age.
            builder.InsertTextInput("Age", TextFormFieldType.Regular, "", "", 3);

            // Save the template.
            templateDoc.Save(templatePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a CSV file with sample data if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(csvPath))
        {
            // Header: Name,Email,Age
            // One data row.
            string csvContent = "Name,Email,Age\r\nJohn Doe,john@example.com,30";
            File.WriteAllText(csvPath, csvContent);
        }

        // -----------------------------------------------------------------
        // 3. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Read CSV data.
        // -----------------------------------------------------------------
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length < 2)
            throw new InvalidOperationException("CSV file does not contain data rows.");

        // Parse header.
        string[] headers = csvLines[0].Split(',');
        // Parse first data row (for simplicity we use only the first row).
        string[] values = csvLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header and data column counts do not match.");

        // Build a dictionary of field name -> value.
        var fieldValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++)
        {
            string key = headers[i].Trim();
            string val = values[i].Trim();
            if (!string.IsNullOrEmpty(key))
                fieldValues[key] = val;
        }

        // -----------------------------------------------------------------
        // 5. Populate form fields in the document.
        // -----------------------------------------------------------------
        foreach (KeyValuePair<string, string> entry in fieldValues)
        {
            // Access the form field by its bookmark/name.
            FormField formField = doc.Range.FormFields[entry.Key];
            if (formField == null)
                throw new InvalidOperationException($"Form field '{entry.Key}' not found in the template.");

            // For text input fields, set the result.
            // Using SetTextInputValue applies any format; Result works as well for plain text.
            formField.SetTextInputValue(entry.Value);
        }

        // -----------------------------------------------------------------
        // 6. Save the populated document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
