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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple numbered list.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Enable restarting the list at each new section – this works only with
        // OOXML compliance higher than Ecma376.
        list.IsRestartAtEachSection = true;

        // Prepare save options with a higher compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional; // higher than Ecma376

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomList.docx");

        // Save the document using the specified options.
        doc.Save(outputPath, saveOptions);

        // Load the saved document to verify that the custom list setting persisted.
        Document loadedDoc = new Document(outputPath);
        bool isRestart = loadedDoc.Lists[0].IsRestartAtEachSection;

        // Output the verification result.
        Console.WriteLine($"IsRestartAtEachSection persisted: {isRestart}");
    }
}
