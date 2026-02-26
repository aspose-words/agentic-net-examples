using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using System.Drawing;

class RenderDocumentWithAlternateHyperlinkDescriptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Below are two hyperlinks with alternate descriptions (ScreenTips):");

        // Insert the first hyperlink.
        // The InsertHyperlink method returns a generic Field; cast it to FieldHyperlink to access hyperlink‑specific properties.
        Field hyperlink1 = builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);
        FieldHyperlink fieldHyperlink1 = (FieldHyperlink)hyperlink1;
        // Set an alternate description that appears as a tooltip when the user hovers over the link.
        fieldHyperlink1.ScreenTip = "Aspose – .NET components for document processing";

        // Insert a line break between the links.
        builder.Writeln();

        // Insert the second hyperlink that points to a bookmark inside the same document.
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("Bookmark target text.");
        builder.EndBookmark("TargetBookmark");

        // Insert a hyperlink to the bookmark.
        Field hyperlink2 = builder.InsertHyperlink("Go to Bookmark", "TargetBookmark", true);
        FieldHyperlink fieldHyperlink2 = (FieldHyperlink)hyperlink2;
        fieldHyperlink2.ScreenTip = "Jump to the bookmarked section within this document";

        // Optionally, demonstrate an image with alternative text (not a hyperlink) for completeness.
        // Insert a shape (image) and set its AlternativeText property.
        Shape shape = builder.InsertShape(ShapeType.Image, 100, 100);
        shape.ImageData.SetImage("https://www.aspose.com/images/aspose-logo.png");
        shape.AlternativeText = "Aspose logo – displayed when the image cannot be loaded";

        // Update all fields so that the hyperlink results are calculated.
        doc.UpdateFields();

        // Save the document to a .docx file.
        doc.Save("HyperlinksWithAlternateDescriptions.docx");
    }
}
