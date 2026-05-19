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
        builder.InsertComboBox("ComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this: ");
        builder.InsertCheckBox("CheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Placeholder", 50);

        // Retrieve the count of form fields in the document's range.
        int formFieldCount = doc.Range.FormFields.Count;

        // Output the count.
        Console.WriteLine($"Number of form fields in the document: {formFieldCount}");

        // Save the document (optional, demonstrates that the document was created).
        doc.Save("FormFields.docx");
    }
}
