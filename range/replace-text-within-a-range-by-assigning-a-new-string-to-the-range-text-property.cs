using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial text.
        builder.Writeln("Hello World!");

        // Replace the word "World" with "Aspose" in the whole document range.
        int replacements = doc.Range.Replace("World", "Aspose");

        // Save the modified document.
        const string outputFile = "Modified.docx";
        doc.Save(outputFile);

        // Output information about the operation.
        Console.WriteLine($"Replacements performed: {replacements}");
        Console.WriteLine($"Document saved to: {outputFile}");
    }
}
