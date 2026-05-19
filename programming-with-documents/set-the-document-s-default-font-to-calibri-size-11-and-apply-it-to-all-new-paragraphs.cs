using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the document-wide default font to Calibri, size 11.
        doc.Styles.DefaultFont.Name = "Calibri";
        doc.Styles.DefaultFont.Size = 11;

        // Add paragraphs; they will inherit the default font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph uses the default Calibri 11 font.");
        builder.Writeln("Second paragraph also uses the default Calibri 11 font.");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFont.docx");
        doc.Save(outputPath);
    }
}
