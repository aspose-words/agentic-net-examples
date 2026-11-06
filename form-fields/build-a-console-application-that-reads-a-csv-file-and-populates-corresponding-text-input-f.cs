using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;   // Required for TextFormFieldType

public class Program
{
    public static void Main()
    {
        // Paths for the template, CSV data, and output document.
        const string templatePath = "Template.docx";
        const string csvPath = "data.csv";
        const string outputPath = "Output.docx";

        // Create a template document with text input form fields if it does not exist.
        if (!File.Exists(templatePath))
        {
            CreateTemplateDocument(templatePath);
        }

        // Create a sample CSV file with headers matching the form field names if it does not exist.
        if (!File.Exists(csvPath))
        {
            CreateSampleCsv(csvPath);
        }

        // Load the template document.
        Document doc = new Document(templatePath);
        FormFieldCollection formFields = doc.Range.FormFields;

        // Read CSV content.
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length < 2)
        {
            throw new Exception("CSV file must contain a header line and at least one data line.");
        }

        // Parse header and first data row.
        string[] headers = csvLines[0].Split(',');
        string[] values = csvLines[1].Split(',');

        // Populate each form field with the corresponding CSV value.
        for (int i = 0; i < headers.Length; i++)
        {
            string fieldName = headers[i].Trim();
            string fieldValue = i < values.Length ? values[i].Trim() : string.Empty;

            // Retrieve the form field by name.
            FormField field = formFields[fieldName];
            if (field == null)
            {
                throw new Exception($"Form field '{fieldName}' not found in the document.");
            }

            // Update the text input field's result.
            field.Result = fieldValue;
        }

        // Save the populated document.
        doc.Save(outputPath);
    }

    // Creates a simple template document with three text input form fields.
    private static void CreateTemplateDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Please fill in the following information:");
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "John", 0);
        builder.Writeln();
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "Doe", 0);
        builder.Writeln();
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "john.doe@example.com", 0);
        builder.Writeln();

        doc.Save(path);
    }

    // Creates a sample CSV file with headers matching the form field names.
    private static void CreateSampleCsv(string path)
    {
        string[] lines =
        {
            "FirstName,LastName,Email",
            "Alice,Smith,alice.smith@example.com"
        };
        File.WriteAllLines(path, lines);
    }
}
