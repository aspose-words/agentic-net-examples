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
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
            builder.InsertParagraph();

            // Insert a check box form field.
            builder.Write("Accept terms: ");
            FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
            builder.InsertParagraph();

            // Insert a text input form field.
            builder.Write("Enter name: ");
            FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
            builder.InsertParagraph();

            // Save the document with form fields (optional, just to demonstrate the before state).
            doc.Save("DocumentWithFormFields.docx");

            // Delete all form fields by iterating over the FormFields collection and removing each field.
            FormFieldCollection formFields = doc.Range.FormFields;
            for (int i = formFields.Count - 1; i >= 0; i--)
            {
                // Remove the complete form field.
                formFields[i].RemoveField();
            }

            // Save the resulting document without form fields.
            doc.Save("DocumentWithoutFormFields.docx");
        }
    }
}
