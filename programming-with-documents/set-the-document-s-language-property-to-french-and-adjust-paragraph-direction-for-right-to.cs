using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default language to French (Locale ID 1036).
        builder.Font.LocaleId = 1036; // French

        // Add a paragraph with French text.
        builder.Writeln("Ceci est un texte en français.");

        // Adjust paragraph direction for right-to-left scripts.
        builder.CurrentParagraph.ParagraphFormat.Bidi = true;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved and can be reopened.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
        }
    }
}
