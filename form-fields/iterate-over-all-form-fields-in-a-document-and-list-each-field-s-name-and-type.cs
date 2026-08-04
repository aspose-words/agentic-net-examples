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

        // Insert a few different types of form fields so that the document is not empty.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);

        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("TermsCheck", false, 15);

        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a fruit: ");
        string[] fruits = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitChoice", fruits, 0);

        // Save the document (required by the rules when modifying form fields).
        doc.Save("FormFieldsIterate.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Ensure that at least one form field exists.
        if (formFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Iterate over each form field and output its name and type.
        foreach (FormField field in formFields)
        {
            // Guard against null entries (should not happen, but follows nullable safety rules).
            if (field != null)
            {
                Console.WriteLine($"Field Name: {field.Name}, Field Type: {field.Type}");
            }
        }
    }
}
