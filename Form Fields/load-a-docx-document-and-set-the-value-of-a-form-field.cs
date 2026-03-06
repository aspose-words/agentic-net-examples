using System;
using Aspose.Words;
using Aspose.Words.Fields;

class SetFormFieldValue
{
    static void Main()
    {
        // Path to the folder that contains the input document.
        string docsPath = @"C:\Docs\";

        // Load the existing DOCX file.
        Document doc = new Document(docsPath + "input.docx");

        // Retrieve the form field by its bookmark name (the name given when the field was inserted).
        // Replace "MyTextInput" with the actual name of your form field.
        FormField formField = doc.Range.FormFields["MyTextInput"];

        // Ensure the field exists and is a text input field before setting its value.
        if (formField != null && formField.Type == FieldType.FieldFormTextInput)
        {
            // Set the new value for the text input form field.
            formField.SetTextInputValue("New value for the field");
        }

        // Save the modified document.
        doc.Save(docsPath + "output.docx");
    }
}
