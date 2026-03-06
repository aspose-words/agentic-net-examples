using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template.
        Document doc = new Document("Template.pdf");

        // Create XPS save options (default settings are sufficient for field evaluation).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the document as XPS; fields are evaluated automatically.
        doc.Save("Result.xps", xpsOptions);
    }
}
