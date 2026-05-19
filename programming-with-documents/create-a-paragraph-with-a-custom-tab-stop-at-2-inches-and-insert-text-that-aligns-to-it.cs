using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom tab stop at 2 inches (2 * 72 points = 144 points).
        // Left alignment and no leader line are used.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph where text before the tab is left‑aligned
        // and text after the tab aligns to the custom tab stop.
        builder.Writeln("Item" + ControlChar.Tab + "Price");

        // Save the document to the local file system.
        doc.Save("CustomTabStop.docx");
    }
}
