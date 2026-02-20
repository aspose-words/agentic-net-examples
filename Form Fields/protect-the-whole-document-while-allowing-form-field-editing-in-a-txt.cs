using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class ProtectTxtWithFormFields
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the following information:");
        // Insert a regular text input form field.
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter name here", 30);

        // Protect the entire document but allow editing of form fields only.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document as a plain‑text file.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        doc.Save("ProtectedFormFields.txt", saveOptions);
    }
}
