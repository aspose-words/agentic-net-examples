using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Path for the sample and result documents.
        string samplePath = "Sample.docx";
        string resultPath = "Result.docx";

        // 1. Create a sample document containing placeholder text.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Welcome to {{Company}}. We hope you enjoy your stay.");
        // Save the sample document.
        sampleDoc.Save(samplePath);

        // 2. Simulate retrieving dynamic data from a database.
        // In a real scenario, replace this with actual DB access code.
        Dictionary<string, string> dbData = new Dictionary<string, string>
        {
            { "Name", "Alice Johnson" },
            { "Company", "Contoso Ltd." }
        };

        // 3. Load the document to be processed.
        Document doc = new Document(samplePath);

        // 4. Replace each placeholder with the corresponding value.
        foreach (var pair in dbData)
        {
            string placeholder = $"{{{{{pair.Key}}}}}"; // e.g., {{Name}}
            // Use Range.Replace for simple text replacement.
            doc.Range.Replace(placeholder, pair.Value, new FindReplaceOptions(FindReplaceDirection.Forward));
        }

        // 5. Save the updated document.
        doc.Save(resultPath);

        // Optional: Verify replacement by outputting the document's text to console.
        // This does not require user interaction.
        Console.WriteLine("Replacement completed. Resulting document text:");
        Console.WriteLine(doc.Range.Text);
    }
}
