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

        // Insert a text input form field.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (unused here)
            "John Doe",                      // placeholder text
            50);                             // maximum length
        // Set an explicit result value.
        textField.Result = "Alice";

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "AcceptTerms",   // field name
            false,           // default unchecked
            50);             // size in points
        // Mark the checkbox as checked.
        checkBox.Checked = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field.
        builder.Write("Select option: ");
        FormField comboBox = builder.InsertComboBox(
            "Options",                     // field name
            new[] { "Option1", "Option2", "Option3" }, // items
            0);                            // default selected index
        // Change the selected item to the third entry.
        comboBox.DropDownSelectedIndex = 2;

        // Save the document so that the form fields are persisted.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFields.docx");
        doc.Save(outputPath);

        // Retrieve the collection of form fields from the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields == null || formFields.Count == 0)
        {
            Console.WriteLine("No form fields were found in the document.");
            return;
        }

        // Iterate through each form field and log its Result value.
        foreach (FormField field in formFields)
        {
            // Guard against a possible null entry.
            if (field == null)
                continue;

            // For text fields the Result holds the entered text.
            // For checkboxes the Result is "1" (checked) or "0" (unchecked).
            // For dropdowns the Result is the selected item text.
            string result = field.Result ?? string.Empty;

            Console.WriteLine($"Field Name: {field.Name}, Type: {field.Type}, Result: \"{result}\"");
        }
    }
}
