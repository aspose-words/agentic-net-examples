using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Settings;

class ProtectHtmlFormFields
{
    static void Main()
    {
        // Load the HTML document. The HtmlLoadOptions can be omitted if default behavior is sufficient.
        var loadOptions = new HtmlLoadOptions();
        Document doc = new Document("input.html", loadOptions);

        // Apply protection that allows only form fields to be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document back to HTML format.
        doc.Save("output.html");
    }
}
