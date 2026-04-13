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
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            "TextField",                     // name
            TextFormFieldType.Regular,       // type
            "",                              // format (none)
            "John Doe",                      // default text
            50);                             // max length

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox",                      // name
            false,                           // default unchecked
            50);                             // size in points

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox(
            "ComboBox",                      // name
            items,                           // items
            0);                              // default selected index

        // Ensure at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created.");
        }

        // Count each type of form field.
        int textCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        foreach (FormField field in formFields)
        {
            if (field == null) continue; // safety check

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    textCount++;
                    break;
                case FieldType.FieldFormCheckBox:
                    checkBoxCount++;
                    break;
                case FieldType.FieldFormDropDown:
                    dropDownCount++;
                    break;
                default:
                    // Other field types are ignored for this example.
                    break;
            }
        }

        // Output the counts.
        Console.WriteLine($"Text input fields: {textCount}");
        Console.WriteLine($"Checkbox fields: {checkBoxCount}");
        Console.WriteLine($"Dropdown fields: {dropDownCount}");

        // Save the document to disk.
        string outputPath = "FormFieldsCount.docx";
        doc.Save(outputPath);
    }
}
