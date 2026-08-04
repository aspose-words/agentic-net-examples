using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class FormFieldCounter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Name", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("CheckBoxField", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a dropdown (combo box) form field.
        builder.Write("Select a color: ");
        string[] colors = { "Red", "Green", "Blue" };
        builder.InsertComboBox("DropDownField", colors, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document that now contains the form fields.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields_Count.docx");
        doc.Save(outputPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that the document contains at least one form field.
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Counters for each form field type.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        // Iterate through the collection and count each type.
        foreach (FormField field in formFields)
        {
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
            }
        }

        // Output the results.
        Console.WriteLine($"Text input fields: {textInputCount}");
        Console.WriteLine($"Checkbox fields: {checkBoxCount}");
        Console.WriteLine($"Dropdown fields: {dropDownCount}");
    }
}
