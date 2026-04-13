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

        // Insert a checkbox form field with an initial size.
        builder.Write("Original checkbox: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 20);
        // Enable explicit size handling.
        checkBox.IsCheckBoxExactSize = true;

        // Save the document that contains the original checkbox (optional).
        doc.Save("Original.docx");

        // -----------------------------------------------------------------
        // Update the size of the existing checkbox programmatically.
        // -----------------------------------------------------------------

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Retrieve the checkbox by its name and validate its existence.
        FormField existingCheckBox = formFields["MyCheckBox"];
        if (existingCheckBox == null)
            throw new InvalidOperationException("The checkbox form field 'MyCheckBox' was not found.");

        // Ensure the checkbox uses an exact size and set the new size (e.g., 40 points).
        existingCheckBox.IsCheckBoxExactSize = true;
        existingCheckBox.CheckBoxSize = 40.0;

        // Save the modified document.
        doc.Save("Modified.docx");
    }
}
