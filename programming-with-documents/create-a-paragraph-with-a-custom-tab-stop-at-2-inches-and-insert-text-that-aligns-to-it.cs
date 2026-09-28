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

        // 1 inch = 72 points, so 2 inches = 144 points.
        // Add a left-aligned tab stop at 2 inches. Specify TabLeader.None as required.
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(144, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph with a tab character so the text aligns to the custom tab stop.
        builder.Writeln("\tText aligned to 2‑inch tab stop");

        // Save the document to a file.
        string outputPath = "CustomTabStop.docx";
        doc.Save(outputPath);

        // Optional verification that the file was created.
        if (File.Exists(outputPath))
        {
            // Document successfully created.
        }
    }
}
