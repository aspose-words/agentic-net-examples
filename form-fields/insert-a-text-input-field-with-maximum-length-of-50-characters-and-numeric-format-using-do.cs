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

        // Insert a text input form field named "NumericField".
        // Use Number type, no specific format string, default value "0", and maximum length of 50 characters.
        builder.InsertTextInput("NumericField", TextFormFieldType.Number, "", "0", 50);

        // Validate that the field was created.
        FormField field = doc.Range.FormFields["NumericField"];
        if (field == null)
            throw new InvalidOperationException("The expected form field was not found.");

        // Update the field's value to demonstrate writing to it.
        field.Result = "12345";

        // Optional: verify the update.
        if (field.Result != "12345")
            throw new InvalidOperationException("Failed to update the form field value.");

        // Save the document to disk.
        doc.Save("FormFieldExample.docx");
    }
}
