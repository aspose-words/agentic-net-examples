using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font name for any subsequently inserted content.
        builder.Font.Name = "Arial";

        // Validate that the font name was correctly assigned.
        if (builder.Font.Name != "Arial")
        {
            throw new InvalidOperationException("Failed to set the default font name.");
        }

        // Insert a line of text that will use the default font.
        builder.Writeln("This text is formatted with the default font set to Arial.");

        // Define the output file path.
        string outputPath = "DefaultFontExample.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the output file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
