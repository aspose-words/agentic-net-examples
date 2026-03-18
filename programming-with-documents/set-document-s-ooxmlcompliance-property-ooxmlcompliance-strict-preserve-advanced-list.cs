using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // (Optional) Add some content to demonstrate that the document is saved correctly.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample paragraph to test OOXML compliance.");

        // Configure OOXML save options to use the Strict compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        // The strict ISO/IEC 29500:2008 compliance level.
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the document with the specified compliance.
        doc.Save("Output_StrictCompliance.docx", saveOptions);
    }
}
