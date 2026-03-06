using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class RestartListPerSection
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Enable list restart at each section for every list in the document.
        foreach (List list in doc.Lists)
        {
            list.IsRestartAtEachSection = true;
        }

        // Set OOXML compliance higher than Ecma376_2006 so the property is saved.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // Save the modified document.
        doc.Save("Output.docx", saveOptions);
    }
}
