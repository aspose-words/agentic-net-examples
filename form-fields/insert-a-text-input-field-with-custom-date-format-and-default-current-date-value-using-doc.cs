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

        // Define the form field name, date format, and default value (current date).
        string fieldName = "DateField";
        string dateFormat = "MM/dd/yyyy";
        string defaultDate = DateTime.Now.ToString(dateFormat);

        // Insert a text input form field of type Date with the custom format and default value.
        builder.InsertTextInput(fieldName, TextFormFieldType.Date, dateFormat, defaultDate, 0);

        // Validate that the form field was created.
        FormField formField = doc.Range.FormFields[fieldName];
        if (formField == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");
        }

        // Verify that the default value matches the expected current date.
        if (formField.Result != defaultDate)
        {
            throw new InvalidOperationException("The default date value does not match the expected value.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
