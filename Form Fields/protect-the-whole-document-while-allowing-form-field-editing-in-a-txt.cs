using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ProtectTxtWithFormFields
{
    static void Main()
    {
        // Load an existing TXT document (or create a new one if needed)
        Document doc = new Document("input.txt");

        // Apply protection that allows only form field editing.
        // Users will be able to fill in form fields but cannot modify other content.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document back as TXT.
        doc.Save("output.txt");
    }
}
