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

        // Insert a combo box (dropdown) form field.
        builder.Write("Choose a fruit: ");
        builder.InsertComboBox("FruitDropDown", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("AcceptCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("NameTextInput", TextFormFieldType.Regular, "", "Your name here", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document that now contains the form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Counters for each type of form field.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        // Iterate through the collection and count by field type.
        foreach (FormField field in formFields)
        {
            if (field == null) continue; // Safety check.

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
        Console.WriteLine($"Form fields saved to: {outputPath}");
        Console.WriteLine($"Text Input fields: {textInputCount}");
        Console.WriteLine($"Check Box fields: {checkBoxCount}");
        Console.WriteLine($"Drop Down fields: {dropDownCount}");
    }
}
