using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the template, CSV data and the resulting document.
        const string templatePath = "Template.docx";
        const string csvPath = "data.csv";
        const string outputPath = "Output.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template with text input form fields if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("User Information");
            builder.Write("First Name: ");
            builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 50);
            builder.Writeln();

            builder.Write("Last Name: ");
            builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "", 50);
            builder.Writeln();

            builder.Write("Email: ");
            builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 100);
            builder.Writeln();

            templateDoc.Save(templatePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a CSV file with header matching the form field names if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(csvPath))
        {
            // Header row followed by a single data row.
            string csvContent = "FirstName,LastName,Email\nJohn,Doe,john.doe@example.com";
            File.WriteAllText(csvPath, csvContent);
        }

        // -----------------------------------------------------------------
        // 3. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Read and parse the CSV file.
        // -----------------------------------------------------------------
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length < 2)
            throw new InvalidOperationException("CSV file must contain a header line and at least one data line.");

        // Header columns.
        string[] headers = csvLines[0].Split(',');
        // First data row (for simplicity we use only the first row).
        string[] values = csvLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header count does not match value count.");

        // Map each header to its corresponding value.
        var fieldValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < headers.Length; i++)
        {
            string header = headers[i].Trim();
            string value = values[i].Trim();
            if (!string.IsNullOrEmpty(header))
                fieldValues[header] = value;
        }

        // -----------------------------------------------------------------
        // 5. Populate the form fields in the document.
        // -----------------------------------------------------------------
        foreach (KeyValuePair<string, string> kvp in fieldValues)
        {
            // Retrieve the form field by its bookmark/name.
            FormField formField = doc.Range.FormFields[kvp.Key];
            if (formField == null)
                throw new InvalidOperationException($"Form field '{kvp.Key}' was not found in the template.");

            // Update the field's result with the CSV value.
            // Using SetTextInputValue applies any format; Result works as well for plain text.
            formField.SetTextInputValue(kvp.Value);
        }

        // -----------------------------------------------------------------
        // 6. Save the populated document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
