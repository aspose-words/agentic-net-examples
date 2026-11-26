using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to insert as separate paragraphs.
        string[] customTexts = { "First custom paragraph.", "Second custom paragraph.", "Third custom paragraph." };

        // Loop through the data and insert each as a new paragraph using Writeln.
        foreach (string text in customTexts)
        {
            builder.Writeln(text);
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
