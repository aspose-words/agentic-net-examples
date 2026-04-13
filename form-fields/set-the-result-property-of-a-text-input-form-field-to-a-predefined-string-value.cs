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

        // Insert a text input form field with a placeholder value.
        builder.Write("Enter text: ");
        FormField insertedField = builder.InsertTextInput(
            name: "MyTextInput",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Placeholder",
            maxLength: 0);

        // Verify that at least one form field exists in the document.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Locate the text input field by its name.
        FormField targetField = null;
        foreach (FormField ff in formFields)
        {
            if (ff.Name == "MyTextInput")
            {
                targetField = ff;
                break;
            }
        }

        if (targetField == null)
            throw new InvalidOperationException("Form field 'MyTextInput' was not found.");

        // Set the Result property to a predefined string.
        const string newResult = "Hello, Aspose!";
        targetField.Result = newResult;

        // Validate that the value was set correctly.
        if (targetField.Result != newResult)
            throw new InvalidOperationException("Failed to assign the new result to the form field.");

        // Save the modified document.
        doc.Save("ResultFormField.docx");
    }
}
