using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // The IsRestartAtEachSection property is only effective when the OOXML compliance
        // level is higher than Ecma376_2006, so we set it here.
        list.IsRestartAtEachSection = true;

        // Display the compliance level of the newly created document (should be Ecma376_2006).
        Console.WriteLine("Initial document compliance: " + doc.Compliance);

        // Prepare save options with a higher OOXML compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // Save the document using the higher compliance level.
        string filePath = "ListComplianceExample.docx";
        doc.Save(filePath, saveOptions);

        // Load the saved document.
        Document loadedDoc = new Document(filePath);

        // Display the compliance level after loading (should reflect the higher compliance).
        Console.WriteLine("Loaded document compliance: " + loadedDoc.Compliance);
    }
}
