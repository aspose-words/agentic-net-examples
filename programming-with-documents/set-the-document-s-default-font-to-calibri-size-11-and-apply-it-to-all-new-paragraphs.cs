using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Set the default font name and size by modifying the Normal style.
        Style normalStyle = doc.Styles[StyleIdentifier.Normal];
        normalStyle.Font.Name = "Calibri";
        normalStyle.Font.Size = 11;

        // Add some paragraphs; they will inherit the Normal style settings.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph using default Calibri 11.");
        builder.Writeln("Second paragraph also uses default Calibri 11.");

        // Save the document to a file.
        string outputPath = "DefaultFontExample.docx";
        doc.Save(outputPath);
    }
}
