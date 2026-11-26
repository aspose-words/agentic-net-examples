using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph describing the field.
        builder.Writeln("Page number in Roman numerals:");

        // Insert a PAGE field formatted as Roman numerals.
        // The field code "PAGE \\* ROMAN" tells Word to display the page number in Roman numerals.
        builder.InsertField("PAGE \\* ROMAN", "i");

        // Save the document to a temporary location.
        string outputPath = Path.Combine(Path.GetTempPath(), "RomanPageNumber.docx");
        doc.Save(outputPath);

        // Verify that the file was created and can be reloaded.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // No further action needed; the program will exit.
        }
    }
}
