using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a hyperlink.
        builder.Writeln("Please visit the following link:");
        builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com", false);

        // Retrieve the last inserted field (the hyperlink) and set it to open in a new tab.
        FieldHyperlink hyperlink = (FieldHyperlink)doc.Range.Fields[doc.Range.Fields.Count - 1];
        hyperlink.Target = "_blank"; // Open in new browser tab.
        hyperlink.Update();

        // Save the document.
        string outputPath = "Hyperlink.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
