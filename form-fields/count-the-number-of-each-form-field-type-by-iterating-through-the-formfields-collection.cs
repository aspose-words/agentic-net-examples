using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class FormFieldCounter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("CheckBox1", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a dropdown (combo box) form field.
        builder.Write("Select a fruit: ");
        string[] fruits = { "Apple", "Banana", "Cherry" };
        builder.InsertComboBox("DropDown1", fruits, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document (required by the rules).
        doc.Save("FormFieldsCount.docx");

        // Counters for each form field type.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        // Iterate through all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;
        foreach (FormField field in formFields)
        {
            if (field == null) continue; // Safety check.

            // Determine the type of the form field.
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    textInputCount++;
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
        Console.WriteLine($"Text Input Fields: {textInputCount}");
        Console.WriteLine($"Check Box Fields: {checkBoxCount}");
        Console.WriteLine($"DropDown Fields: {dropDownCount}");
    }
}
