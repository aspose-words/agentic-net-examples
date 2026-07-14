using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace FormFieldRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a combo box form field.
            builder.Write("Choose a value: ");
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a check box form field.
            builder.Write("Check this box: ");
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a text input form field.
            builder.Write("Enter text: ");
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Iterate over all form fields in the document's range and output their name and type.
            FormFieldCollection formFields = doc.Range.FormFields;
            foreach (FormField field in formFields)
            {
                Console.WriteLine($"Form field name: {field.Name}, type: {field.Type}");
            }

            // Save the document locally (optional, demonstrates saving workflow).
            doc.Save("FormFields.docx");
        }
    }
}
