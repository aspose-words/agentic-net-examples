using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom tab stop at 2 inches (2 * 72 points = 144 points).
        // The tab stop is left‑aligned with no leader line.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph that uses the tab character to align text to the custom tab stop.
        builder.Writeln("Start" + ControlChar.Tab + "Aligned at 2 inches");

        // Save the resulting document to the current directory.
        doc.Save("CustomTabStop.docx");
    }
}
