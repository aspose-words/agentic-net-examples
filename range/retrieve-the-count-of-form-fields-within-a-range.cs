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
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);

        // Insert a check box form field.
        builder.InsertCheckBox("MyCheckBox", false, 50);

        // Insert a text input form field.
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);

        // Retrieve the count of form fields in the document's range.
        int formFieldCount = doc.Range.FormFields.Count;

        // Output the count to the console.
        Console.WriteLine($"Number of form fields in the document: {formFieldCount}");

        // Save the document (optional, demonstrates that the document was created).
        doc.Save("FormFields.docx");
    }
}
