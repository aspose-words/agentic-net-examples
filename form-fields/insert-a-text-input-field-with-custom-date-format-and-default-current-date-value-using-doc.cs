using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the form field name and custom date format.
        string fieldName = "DateField";
        string dateFormat = "MM/dd/yyyy";

        // Use the current date as the default value, formatted according to the custom format.
        string defaultDate = DateTime.Now.ToString(dateFormat);

        // Insert a text input form field of type Date with the custom format and default value.
        // The maxLength parameter is set to 0, which means no length restriction.
        builder.InsertTextInput(fieldName, TextFormFieldType.Date, dateFormat, defaultDate, 0);

        // Validate that the form field was created and has the expected default value.
        FormField field = doc.Range.FormFields[fieldName];
        if (field == null)
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");

        if (field.Result != defaultDate)
            throw new InvalidOperationException($"Form field default value mismatch. Expected '{defaultDate}', got '{field.Result}'.");

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
