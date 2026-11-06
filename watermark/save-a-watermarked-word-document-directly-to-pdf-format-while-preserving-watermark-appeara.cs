using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a watermark.");

        // Apply a text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document directly to PDF format.
        doc.Save("Watermarked.pdf", SaveFormat.Pdf);
    }
}
