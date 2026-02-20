using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the following form field:");
        // Insert a regular text input form field.
        builder.InsertTextInput("FormField1", TextFormFieldType.Regular, "", "Enter value here", 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the document as a DOT (Word template) file.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save("FormOnlyTemplate.dot", saveOptions);
    }
}
