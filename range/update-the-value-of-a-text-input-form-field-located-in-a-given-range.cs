using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the form field.
        builder.Writeln("Enter your name:");

        // Insert a regular text input form field named "NameField".
        // The field initially shows "John Doe" as placeholder text.
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);

        // Retrieve the form field by its name and update its value.
        FormField nameField = doc.Range.FormFields["NameField"];
        if (nameField != null)
        {
            nameField.SetTextInputValue("Jane Smith");
        }

        // Update any fields in the document (optional but safe).
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("UpdatedFormField.docx");
    }
}
