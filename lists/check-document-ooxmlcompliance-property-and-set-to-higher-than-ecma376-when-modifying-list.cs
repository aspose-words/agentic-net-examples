using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Modify the list definition – enable restarting at each section.
        list.IsRestartAtEachSection = true;

        // Determine the OOXML compliance of the current document.
        OoxmlCompliance currentCompliance = doc.Compliance;

        // Prepare save options. If the document compliance is the default (Ecma376_2006),
        // set a higher compliance level so that the modified list definition is persisted.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        if (currentCompliance == OoxmlCompliance.Ecma376_2006)
        {
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        }

        // Save the document with the specified options.
        string outputPath = "ListCompliance.docx";
        doc.Save(outputPath, saveOptions);

        // Load the saved document to verify the compliance level.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine($"Saved document compliance: {loadedDoc.Compliance}");
    }
}
