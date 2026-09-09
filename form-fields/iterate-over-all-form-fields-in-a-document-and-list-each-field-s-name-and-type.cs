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
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a fruit: ");
        string[] items = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitChoice", items, 0);

        // Save the document (required by the rules).
        doc.Save("FormFields.docx");

        // Iterate over all form fields and list their name and type.
        FormFieldCollection formFields = doc.Range.FormFields;
        foreach (FormField field in formFields)
        {
            // Field.Type returns a FieldType enum value.
            Console.WriteLine($"{field.Name}: {field.Type}");
        }
    }
}
