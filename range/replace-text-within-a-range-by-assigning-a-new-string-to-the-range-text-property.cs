using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some initial text.
        builder.Writeln("Hello Aspose.Words!");

        // Replace the target text within the document's range.
        // This uses the native Range.Replace method as required.
        int replacements = doc.Range.Replace("Aspose.Words", "World");

        // Save the modified document to the local file system.
        const string outputFile = "Output.docx";
        doc.Save(outputFile);

        // Output verification information.
        Console.WriteLine($"Replacements performed: {replacements}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
