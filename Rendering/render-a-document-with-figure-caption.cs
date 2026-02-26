using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FigureCaptionExample
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image that will act as the figure.
        // Replace the path with a valid image file on your system.
        builder.InsertImage(@"C:\Images\SampleFigure.png");

        // Insert a paragraph break after the image.
        builder.Writeln();

        // Build the caption: "Figure 1: Sample figure description"
        builder.Write("Figure ");
        // Insert a SEQ field that numbers figures. The field result will be updated later.
        builder.InsertField(" SEQ Figure \\* ARABIC ", "");
        builder.Write(": Sample figure description.");

        // Insert another paragraph break to separate the caption from the rest of the content.
        builder.Writeln();

        // Insert a Table of Figures field. This field will collect all SEQ Figure entries.
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
        // Specify that the TOC should use the "Figure" sequence identifier.
        toc.TableOfFiguresLabel = "Figure";

        // Update all fields in the document so that the figure number and the table of figures are correct.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save(@"C:\Output\FigureWithCaption.docx");
    }
}
