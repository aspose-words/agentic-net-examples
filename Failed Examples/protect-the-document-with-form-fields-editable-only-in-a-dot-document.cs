// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class FormTemplateCreator
{
    static void Main()
    {
        // Path where the template will be saved.
        string outputPath = "FormTemplate.dot";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the form below:");
        // Insert a regular text input form field.
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter name", 0);

        // Protect the whole document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the document as a DOT (Word template) file.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Dot);
        doc.Save(outputPath, saveOptions);
    }
}
