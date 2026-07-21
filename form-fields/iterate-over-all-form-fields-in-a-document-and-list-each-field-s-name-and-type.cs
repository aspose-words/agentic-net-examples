using System;
using System.IO;
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
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field.
        builder.Write("Select a country: ");
        string[] countries = { "USA", "Canada", "Mexico" };
        builder.InsertComboBox("CountryField", countries, 0);

        // Save the document that now contains form fields.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields.docx");
        doc.Save(outputPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate over each form field and output its name and type.
        foreach (FormField field in formFields)
        {
            // Field.Type is an enum; ToString() gives a readable name.
            Console.WriteLine($"Field Name: {field.Name}, Field Type: {field.Type}");
        }
    }
}
