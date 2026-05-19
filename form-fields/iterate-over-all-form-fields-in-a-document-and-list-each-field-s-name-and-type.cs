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

        // Insert a few different types of legacy form fields.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);

        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms? ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 15);

        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a fruit: ");
        FormField comboBox = builder.InsertComboBox("FruitChoice", new[] { "Apple", "Banana", "Cherry" }, 0);

        // Retrieve the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that the document contains at least one form field.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate over the collection and output each field's name and type.
        foreach (FormField field in formFields)
        {
            // Guard against a possible null entry (should not happen, but safe).
            if (field != null)
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }

        // Save the document (required by the rules, even though we only read fields).
        doc.Save("FormFields.docx");
    }
}
