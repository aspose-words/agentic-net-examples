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

        // Insert a legacy text input form field.
        // Parameters: name, type, format, default text, maxLength.
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Default", 0);

        // Ensure that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the document.");
        }

        // Access the first form field by index.
        FormField formField = doc.Range.FormFields[0];
        if (formField == null)
        {
            throw new InvalidOperationException("Form field at index 0 is null.");
        }

        // Modify the Result property of the text input field.
        string newValue = "Updated value";
        formField.Result = newValue;

        // Validate that the value was updated.
        if (formField.Result != newValue)
        {
            throw new InvalidOperationException("Failed to update the form field value.");
        }

        // Save the modified document.
        doc.Save("FormFieldModified.docx");
    }
}
