using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        const string csvPath = "data.csv";
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple CSV file that will be used as the data source.
        // -----------------------------------------------------------------
        File.WriteAllText(csvPath,
            "FirstName,LastName,Email\n" +
            "John,Doe,john.doe@example.com");

        // ---------------------------------------------------------------
        // 2. Build a Word template containing text input form fields.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Write("First Name: ");
        builder.InsertTextInput("FirstName", TextFormFieldType.Regular, "", "", 50);
        builder.Writeln();

        builder.Write("Last Name: ");
        builder.InsertTextInput("LastName", TextFormFieldType.Regular, "", "", 50);
        builder.Writeln();

        builder.Write("Email: ");
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 100);
        builder.Writeln();

        // Save the template for later loading.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and populate its form fields from CSV.
        // ---------------------------------------------------------------
        Document doc = new Document(templatePath);
        FormFieldCollection formFields = doc.Range.FormFields;

        // Read CSV content.
        string[] csvLines = File.ReadAllLines(csvPath);
        if (csvLines.Length < 2)
            throw new InvalidOperationException("CSV file must contain a header and at least one data row.");

        // Header defines the mapping between CSV columns and form field names.
        string[] headers = csvLines[0].Split(',');

        // Use the first data row for this example.
        string[] values = csvLines[1].Split(',');

        if (headers.Length != values.Length)
            throw new InvalidOperationException("CSV header and data column counts do not match.");

        // Populate each form field.
        for (int i = 0; i < headers.Length; i++)
        {
            string fieldName = headers[i];
            string fieldValue = values[i];

            // Validate field existence.
            FormField field = formFields[fieldName];
            if (field == null)
                throw new InvalidOperationException($"Form field '{fieldName}' not found in the document.");

            // Set the value. Using SetTextInputValue applies any format (none in this case).
            field.SetTextInputValue(fieldValue);
        }

        // ---------------------------------------------------------------
        // 4. Save the populated document.
        // ---------------------------------------------------------------
        doc.Save(outputPath);
    }
}
