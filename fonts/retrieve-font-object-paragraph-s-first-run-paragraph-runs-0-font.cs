using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Remove any default nodes and set up a minimal structure.
        doc.RemoveAllChildren();
        Section section = new Section(doc);
        doc.AppendChild(section);
        Body body = new Body(doc);
        section.AppendChild(body);
        Paragraph para = new Paragraph(doc);
        body.AppendChild(para);

        // Add a run of text to the paragraph.
        Run run = new Run(doc, "Hello World!");
        para.AppendChild(run);

        // Retrieve the Font object from the paragraph's first run.
        Font font = para.Runs[0].Font;

        // Modify some font properties as an example.
        font.Name = "Courier New";
        font.Size = 24;
        font.Color = Color.Blue;

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
