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

        // Insert a combo box form field.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document with form fields (optional, shows original state).
        doc.Save("Original.docx");

        // Delete all form fields by iterating over the FormFields collection.
        FormFieldCollection formFields = doc.Range.FormFields;
        for (int i = formFields.Count - 1; i >= 0; i--)
        {
            // Remove the complete form field.
            formFields[i].RemoveField();
        }

        // Save the document after removal of all form fields.
        doc.Save("NoFormFields.docx");
    }
}
