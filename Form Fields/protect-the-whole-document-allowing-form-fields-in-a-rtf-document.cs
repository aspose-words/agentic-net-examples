using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class ProtectRtfWithFormFields
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the following form field:");
        // Insert a regular text input form field.
        builder.InsertTextInput("FormField1", TextFormFieldType.Regular, "", "Enter text here", 0);

        // Protect the entire document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Prepare RTF save options (optional: customize if needed).
        RtfSaveOptions saveOptions = new RtfSaveOptions();
        // Example: keep default settings; you can modify properties such as ExportCompactSize here.

        // Save the protected document as RTF.
        doc.Save("ProtectedFormFields.rtf", saveOptions);
    }
}
