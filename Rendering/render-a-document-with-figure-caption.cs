using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class FigureCaptionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image (replace with a valid path to an image file on your system).
        // The InsertImage method returns the Shape that represents the picture.
        Shape picture = builder.InsertImage("C:\\Images\\SampleImage.jpg");

        // Add a paragraph for the figure caption.
        builder.Writeln(); // Move to a new line after the picture.

        // Write the static part of the caption.
        builder.Write("Figure ");

        // Insert a SEQ field that will generate the figure number.
        // The field code "SEQ Figure \\* ARABIC" creates a sequence named "Figure".
        // The second argument is an empty placeholder for the field result.
        builder.InsertField("SEQ Figure \\* ARABIC", "");

        // Write the rest of the caption text.
        builder.Write(": Sample figure caption describing the image.");

        // Insert a page break before the Table of Figures.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a Table of Figures field.
        // The field type FieldTOC is used for both TOC and Table of Figures.
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);

        // Configure the TOC to build a Table of Figures for the "Figure" sequence.
        toc.TableOfFiguresLabel = "Figure";

        // Optionally, set the captionless label if you need a table without the label/number.
        // toc.CaptionlessTableOfFiguresLabel = "Figure";

        // Update all fields in the document so that the figure number and the Table of Figures are populated.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("FigureCaption.docx");
    }
}
