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

        // Set the default font name that will be applied to all newly inserted text.
        builder.Font.Name = "Arial";

        // Validate that the font name was correctly assigned.
        if (builder.Font.Name != "Arial")
            throw new InvalidOperationException("Failed to set the default font name.");

        // Insert some text; it will use the default font set above.
        builder.Writeln("This text is formatted with the default font 'Arial'.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
