using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for TextFormFieldType

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please fill in the form field below:");

        // Insert a plain text form field that the user can edit.
        builder.InsertTextInput("FormField1", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Protect the document so that only form fields are editable.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the document as a DOT (Word template) file.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save("FormTemplate.dot", saveOptions);
    }
}
