using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Clear any existing tab stops and add a custom tab stop at 2 inches (144 points).
        // The third argument specifies the tab leader; use None for a plain tab.
        builder.ParagraphFormat.TabStops.Clear();
        builder.ParagraphFormat.TabStops.Add(144, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph with a tab character so the text aligns to the custom tab stop.
        builder.Writeln("\tThis text aligns to the custom tab stop at 2 inches.");

        // Save the document to disk.
        string outputPath = "CustomTabStop.docx";
        doc.Save(outputPath);
    }
}
