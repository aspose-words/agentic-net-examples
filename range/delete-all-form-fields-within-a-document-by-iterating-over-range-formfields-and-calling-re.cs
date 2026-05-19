using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace DeleteFormFieldsExample
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
            builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a check box form field.
            builder.Write("Accept terms: ");
            builder.InsertCheckBox("MyCheckBox", false, 50);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Insert a text input form field.
            builder.Write("Enter name: ");
            builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
            builder.InsertBreak(BreakType.ParagraphBreak);

            // Save the document with form fields (optional, just for reference).
            doc.Save("FormFields.docx");

            // Remove all form fields from the document.
            // The FormFieldCollection.Clear method removes every form field and updates the document.
            doc.Range.FormFields.Clear();

            // Save the document after all form fields have been removed.
            doc.Save("FormFields_NoFields.docx");
        }
    }
}
