using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box form field.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);

        // Retrieve the form fields collection from the document's range.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Get the count of form fields.
        int count = formFields.Count;

        // Output the count.
        Console.WriteLine($"Number of form fields in the document range: {count}");

        // Save the document (optional, demonstrates that the document was created).
        doc.Save("FormFields.docx");
    }
}
