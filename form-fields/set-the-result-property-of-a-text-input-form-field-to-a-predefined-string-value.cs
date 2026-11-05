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

        // Insert a text input form field with a name and a default value.
        string fieldName = "MyTextField";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Default", 0);

        // Save the document that contains the form field.
        string initialPath = "FormFieldDemo.docx";
        doc.Save(initialPath);

        // Retrieve the form field from the document.
        FormField? formField = doc.Range.FormFields[fieldName];
        if (formField == null)
            throw new InvalidOperationException($"Form field '{fieldName}' was not found.");

        // Set the Result property to a predefined string.
        string predefinedValue = "Predefined string value";
        formField.Result = predefinedValue;

        // Validate that the value was set correctly.
        if (formField.Result != predefinedValue)
            throw new InvalidOperationException("Failed to set the form field result.");

        // Save the updated document.
        string outputPath = "FormFieldResult.docx";
        doc.Save(outputPath);
    }
}
