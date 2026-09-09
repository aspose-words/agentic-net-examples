using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];

        // Display the OOXML compliance of the newly created document (should be Ecma376_2006).
        Console.WriteLine("Initial document compliance: " + doc.Compliance);

        // Modify a list definition that requires a higher OOXML compliance level.
        // For example, enable restarting the list at each section.
        list.IsRestartAtEachSection = true;

        // Prepare save options with a compliance level higher than Ecma376.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document using the higher compliance settings.
        doc.Save("ModifiedList.docx", saveOptions);

        // Load the saved document to verify the compliance level.
        Document loadedDoc = new Document("ModifiedList.docx");
        Console.WriteLine("Saved document compliance: " + loadedDoc.Compliance);
    }
}
