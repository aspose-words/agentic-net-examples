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
        // Name: "MyTextField", type: regular text, default text: "Placeholder", max length: 50.
        FormField textField = builder.InsertTextInput(
            "MyTextField",
            TextFormFieldType.Regular,
            "",
            "Placeholder",
            50);

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created in the document.");
        }

        // Access the first form field by index (zero‑based) and modify its Result.
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex == null)
        {
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");
        }

        // Set a new value for the text input field.
        fieldByIndex.Result = "New value set by code";

        // Optional: verify that the value was updated.
        if (fieldByIndex.Result != "New value set by code")
        {
            throw new InvalidOperationException("Failed to update the form field result.");
        }

        // Save the modified document.
        string outputPath = "ModifiedFormFields.docx";
        doc.Save(outputPath);
    }
}
