using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box (drop‑down) form field.
        builder.Write("Choose a fruit: ");
        builder.InsertComboBox("FruitDropDown", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("AcceptCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter name: ");
        builder.InsertTextInput("NameTextInput", TextFormFieldType.Regular, "", "Your name", 30);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document (required by the rules).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields_Count.docx");
        doc.Save(outputPath);

        // Counters for each form field type.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        // Iterate through the FormFields collection and count each type.
        FormFieldCollection formFields = doc.Range.FormFields;
        foreach (FormField field in formFields)
        {
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
                    dropDownCount++;
                    break;
            }
        }

        // Output the counts.
        Console.WriteLine($"Text input fields: {textInputCount}");
        Console.WriteLine($"Check box fields: {checkBoxCount}");
        Console.WriteLine($"Drop‑down fields: {dropDownCount}");
    }
}
