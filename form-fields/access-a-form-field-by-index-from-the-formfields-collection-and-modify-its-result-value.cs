using System;
using System.IO;
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
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (unused here)
            "John Doe",                      // placeholder text
            50);                             // maximum length

        // Retrieve the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Access the first form field by its zero‑based index.
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex == null)
            throw new InvalidOperationException("Form field at index 0 could not be retrieved.");

        // Modify the Result property of the form field.
        fieldByIndex.Result = "Jane Smith";

        // Verify that the value was updated correctly.
        if (fieldByIndex.Result != "Jane Smith")
            throw new InvalidOperationException("Failed to update the form field result.");

        // Save the modified document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFieldResultUpdated.docx");
        doc.Save(outputPath);
    }
}
