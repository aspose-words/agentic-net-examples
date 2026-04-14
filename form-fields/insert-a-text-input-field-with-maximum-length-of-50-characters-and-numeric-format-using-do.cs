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

        // Insert a text input form field.
        // - Name: "NumericInput"
        // - Type: Number (allows only numeric input)
        // - Format: "0" (simple numeric format)
        // - Default text: "0"
        // - Maximum length: 50 characters
        FormField textField = builder.InsertTextInput(
            "NumericInput",
            TextFormFieldType.Number,
            "0",
            "0",
            50);

        // Verify that the field was inserted.
        if (textField == null)
            throw new InvalidOperationException("Failed to create the text input form field.");

        // Access the field through the document's form fields collection.
        FormFieldCollection fields = doc.Range.FormFields;
        FormField? retrievedField = fields["NumericInput"];
        if (retrievedField == null)
            throw new InvalidOperationException("The expected form field was not found.");

        // Set a sample numeric value.
        retrievedField.SetTextInputValue(12345);

        // Validate that the value was set correctly.
        if (retrievedField.Result != "12345")
            throw new InvalidOperationException("The form field value was not set as expected.");

        // Save the document.
        doc.Save("FormFieldExample.docx");
    }
}
