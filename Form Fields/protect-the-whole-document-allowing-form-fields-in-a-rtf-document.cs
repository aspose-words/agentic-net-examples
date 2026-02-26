using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for TextFormFieldType

class ProtectRtfWithFormFields
{
    static void Main()
    {
        // Path where the output RTF will be saved.
        string outputPath = @"C:\Output\ProtectedFormFields.rtf";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the form below:");
        // Insert a regular text input form field.
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Protect the entire document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the document as RTF using default RtfSaveOptions.
        RtfSaveOptions saveOptions = new RtfSaveOptions();
        doc.Save(outputPath, saveOptions);
    }
}
