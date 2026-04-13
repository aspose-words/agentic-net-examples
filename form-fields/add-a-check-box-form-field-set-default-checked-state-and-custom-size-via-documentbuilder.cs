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

        // Write some introductory text.
        builder.Writeln("Please tick the checkbox below:");

        // Insert a checkbox form field with a name, default checked state, current checked state, and custom size (30 points).
        // The overload with four parameters allows us to set the default value explicitly.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", defaultValue: true, checkedValue: true, size: 30);

        // Enable exact size so that the custom size is applied.
        checkBox.IsCheckBoxExactSize = true;

        // Optional: add a paragraph break after the checkbox.
        builder.InsertParagraph();

        // Validate that the form field was added correctly.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Retrieve the inserted checkbox by name.
        FormField retrieved = formFields["MyCheckBox"];
        if (retrieved == null)
            throw new InvalidOperationException("The checkbox form field 'MyCheckBox' was not found.");

        // Verify the default and current checked states.
        if (!retrieved.Default || !retrieved.Checked)
            throw new InvalidOperationException("The checkbox default or checked state is not set as expected.");

        // Verify the custom size.
        if (retrieved.CheckBoxSize != 30)
            throw new InvalidOperationException("The checkbox size was not set correctly.");

        // Save the document to disk.
        doc.Save("CheckBoxFormField.docx");
    }
}
