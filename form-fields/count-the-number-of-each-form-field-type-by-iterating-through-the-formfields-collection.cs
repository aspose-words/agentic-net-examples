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

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("CheckBoxField", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (drop‑down) form field.
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        builder.InsertComboBox("ComboBoxField", items, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document (required by the rules).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFields_Count.docx");
        doc.Save(outputPath);

        // Count each type of form field.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int comboBoxCount = 0;

        FormFieldCollection formFields = doc.Range.FormFields;
        foreach (FormField field in formFields)
        {
            // Guard against possible null entries.
            if (field == null) continue;

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    textInputCount++;
                    break;
                case FieldType.FieldFormCheckBox:
                    checkBoxCount++;
                    break;
                case FieldType.FieldFormDropDown:
                    comboBoxCount++;
                    break;
            }
        }

        // Output the results.
        Console.WriteLine($"Text Input fields: {textInputCount}");
        Console.WriteLine($"Check Box fields: {checkBoxCount}");
        Console.WriteLine($"Combo Box (Drop‑Down) fields: {comboBoxCount}");
    }
}
