using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting the list at each section (requires OOXML compliance higher than Ecma376).
        list.IsRestartAtEachSection = true;

        // Verify the current compliance of the in‑memory document (should be Ecma376_2006).
        Console.WriteLine("Document compliance before saving: " + doc.Compliance);

        // Prepare save options with a higher compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            // Use a compliance level newer than Ecma376 to preserve the list setting.
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // Define the output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ListRestart.docx");

        // Save the document with the specified compliance.
        doc.Save(outputPath, saveOptions);

        // Load the saved document to confirm the settings were persisted.
        Document loadedDoc = new Document(outputPath);

        // Output the compliance of the loaded document.
        Console.WriteLine("Document compliance after loading: " + loadedDoc.Compliance);

        // Output the IsRestartAtEachSection value of the first list.
        Console.WriteLine("IsRestartAtEachSection: " + loadedDoc.Lists[0].IsRestartAtEachSection);
    }
}
