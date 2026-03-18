using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some explanatory text.
        builder.Writeln("Please fill in the form below:");

        // Insert a regular text input form field.
        // The user will be able to edit this field after the document is protected.
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Protect the whole document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save("ProtectedFormFields.docx");
    }
}
