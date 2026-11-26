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

        // Insert a hyperlink field and cast it to FieldHyperlink to set the target.
        FieldHyperlink hyperlink = (FieldHyperlink)builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);
        // Configure the hyperlink to open in a new browser tab.
        hyperlink.Target = "_blank";

        // Add a line break after the hyperlink.
        builder.Writeln();

        // Save the document.
        string outputPath = "Hyperlink.docx";
        doc.Save(outputPath);
    }
}
