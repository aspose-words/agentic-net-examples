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

        // Add a prompt and a text input form field.
        builder.Writeln("Please fill in the form below:");
        builder.Write("Name: ");
        FormField nameField = builder.InsertTextInput(
            "NameField",                     // Field name
            TextFormFieldType.Regular,       // Field type
            "",                              // No specific format
            "Enter name here",               // Placeholder text
            0);                              // No length limit

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        const string outputPath = "ProtectedForm.docx";
        doc.Save(outputPath);
    }
}
