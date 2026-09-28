using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Modify the Normal style (applied to new paragraphs) to set the default font name and size.
        Style normalStyle = doc.Styles[StyleIdentifier.Normal];
        normalStyle.Font.Name = "Calibri";
        normalStyle.Font.Size = 11;

        // Use DocumentBuilder to add paragraphs that will inherit the Normal style.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph using Calibri, size 11.");
        builder.Writeln("This is the second paragraph also using Calibri, size 11.");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
