using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph.
        builder.Writeln();

        // Add a tab stop at 2 inches (144 points) with no leader.
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(144, TabAlignment.Left, TabLeader.None);

        // Insert text that aligns to the custom tab stop.
        builder.Write("\tText aligned to a 2‑inch tab stop.");

        // Save the document.
        string outputPath = "TabStopExample.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
