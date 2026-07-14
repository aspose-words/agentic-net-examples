using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists; // Needed for List and ListTemplate types

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the saved document.
        string docPath = Path.Combine(outputDir, "ListWithRestart.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];

        // Modify a list definition that requires a higher OOXML compliance level.
        // For example, enable restarting the list at each section.
        list.IsRestartAtEachSection = true;

        // Check the current compliance of the document (should be Ecma376_2006 for a new document).
        Console.WriteLine($"Document compliance before saving: {doc.Compliance}");

        // Since we modified a property that is only written when the compliance level is higher than Ecma376,
        // create OoxmlSaveOptions with a higher compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document using the specified save options.
        doc.Save(docPath, saveOptions);

        // Load the saved document to verify the compliance level.
        Document loadedDoc = new Document(docPath);
        Console.WriteLine($"Document compliance after loading: {loadedDoc.Compliance}");

        // Output the result of the list property to confirm it was persisted.
        Console.WriteLine($"IsRestartAtEachSection persisted: {loadedDoc.Lists[0].IsRestartAtEachSection}");
    }
}
