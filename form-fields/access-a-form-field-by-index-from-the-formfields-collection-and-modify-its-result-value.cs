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

        // Insert a text input form field with a default placeholder.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            "MyTextInput",                     // field name
            TextFormFieldType.Regular,         // field type
            "",                                // default text (empty)
            "John Doe",                        // placeholder text
            50);                               // maximum length

        // Access the form field by index from the FormFields collection.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Retrieve the first form field (index 0) and ensure it is not null.
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex == null)
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");

        // Modify the Result property of the text input field.
        fieldByIndex.Result = "Jane Smith";

        // Save the modified document.
        doc.Save("FormFieldResultUpdated.docx");
    }
}
