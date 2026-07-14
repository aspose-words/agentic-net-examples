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

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("CheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field.
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        builder.InsertComboBox("DropDown", items, 0);

        // Save the document that now contains form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Iterate over all form fields and list each field's name and type.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        foreach (FormField field in formFields)
        {
            // Field.Type returns a FieldType enum indicating the kind of form field.
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
