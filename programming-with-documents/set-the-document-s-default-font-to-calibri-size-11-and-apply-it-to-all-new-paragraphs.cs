using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the default font name and size via the Normal style.
        Style normalStyle = doc.Styles[StyleIdentifier.Normal];
        normalStyle.Font.Name = "Calibri";
        normalStyle.Font.Size = 11;

        // Add paragraphs to verify the default font is applied.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph using default Calibri 11.");
        builder.Writeln("Second paragraph also using default Calibri 11.");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
