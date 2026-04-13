using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the template, CSV data and the resulting document.
        const string templatePath = "Template.docx";
        const string csvPath = "Data.csv";
        const string outputPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with text input form fields.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a label and a text input field for each column we will use.
        builder.Writeln("First Name:");
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 50);
        builder.Writeln();

        builder.Writeln("Last Name:");
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "", 50);
        builder.Writeln();

        builder.Writeln("Email:");
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 100);
        builder.Writeln();

        // Save the template so it can be loaded later.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a simple CSV file that matches the form field names.
        // -----------------------------------------------------------------
        // Header line corresponds to the form field names.
        // Second line contains the values to populate.
        string[] csvLines =
        {
            "FirstName,LastName,Email",
            "John,Doe,john.doe@example.com"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 3. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Read CSV data.
        // -----------------------------------------------------------------
        string[] allLines = File.ReadAllLines(csvPath);
        if (allLines.Length < 2)
            throw new InvalidOperationException("CSV file must contain at least a header and one data row.");

        // Parse header and first data row.
        string[] headers = allLines[0].Split(',');
        string[] values = allLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header count does not match value count.");

        // -----------------------------------------------------------------
        // 5. Populate form fields with CSV values.
        // -----------------------------------------------------------------
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        for (int i = 0; i < headers.Length; i++)
        {
            string fieldName = headers[i].Trim();
            string fieldValue = values[i].Trim();

            // Access the form field by its bookmark/name.
            FormField field = formFields[fieldName];
            if (field == null)
                throw new InvalidOperationException($"Form field '{fieldName}' not found in the document.");

            // Set the value. Using SetTextInputValue applies any format; Result works as well.
            field.SetTextInputValue(fieldValue);

            // Validate that the value was set correctly.
            if (!string.Equals(field.Result, fieldValue, StringComparison.Ordinal))
                throw new InvalidOperationException($"Failed to set value for field '{fieldName}'.");
        }

        // -----------------------------------------------------------------
        // 6. Save the populated document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
