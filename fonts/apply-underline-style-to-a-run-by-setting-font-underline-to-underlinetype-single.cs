using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the underline style to single.
        builder.Font.Underline = Aspose.Words.Underline.Single;

        // Write a line of underlined text.
        builder.Writeln("This text is underlined with a single line.");

        // Validate that the underline property was set correctly.
        if (builder.Font.Underline != Aspose.Words.Underline.Single)
            throw new InvalidOperationException("Underline style was not applied correctly.");

        // Define the output file name.
        string outputPath = "UnderlinedRun.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved successfully.", outputPath);
    }
}
