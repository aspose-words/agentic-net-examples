using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldCounter
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder to add form fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a text input form field.
            // Parameters: name, type, format, default text, maxLength.
            builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default Text", 0);

            // Insert a checkbox form field.
            // Parameters: name, isChecked, size (in points).
            builder.InsertCheckBox("CheckBox1", false, 0);

            // Insert a dropdown (combo box) form field.
            // Parameters: name, items, selected index.
            builder.InsertComboBox("DropDown1", new string[] { "Option1", "Option2", "Option3" }, 0);

            // Save the document with the created form fields.
            const string outputPath = "FormFields.docx";
            doc.Save(outputPath);

            // Ensure that the document contains at least one form field.
            FormFieldCollection formFields = doc.Range.FormFields;
            if (formFields == null || formFields.Count == 0)
                throw new InvalidOperationException("The document does not contain any form fields.");

            // Counters for each form field type.
            int textInputCount = 0;
            int checkBoxCount = 0;
            int dropDownCount = 0;

            // Iterate through the form fields and count each type.
            foreach (FormField field in formFields)
            {
                if (field == null) continue; // Guard against null entries.

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
            Console.WriteLine($"Text Input Fields: {textInputCount}");
            Console.WriteLine($"Checkbox Fields: {checkBoxCount}");
            Console.WriteLine($"Dropdown Fields: {dropDownCount}");
        }
    }
}
