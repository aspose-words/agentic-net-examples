using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a hyperlink.
        builder.Writeln(); // Ensure we are on a new paragraph.
        builder.InsertHyperlink("Visit Aspose.Words", "https://www.aspose.com", false);

        // Retrieve the inserted hyperlink field and set it to open in a new browser tab.
        // The target "_blank" instructs browsers to open the link in a new tab/window.
        FieldHyperlink hyperlink = (FieldHyperlink)doc.Range.Fields[doc.Range.Fields.Count - 1];
        hyperlink.Target = "_blank";
        hyperlink.Update();

        // Save the document to disk.
        string outputPath = "Hyperlink.docx";
        doc.Save(outputPath);

        // Verify that the file was saved (optional, no output required).
        // The program ends here.
    }
}
