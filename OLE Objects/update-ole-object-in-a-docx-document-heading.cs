// Load an existing DOCX document
var doc = new Aspose.Words.Document("Input.docx");

// Find the first paragraph that contains the heading text "My Heading"
var headingParagraph = doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true)
    .Cast<Aspose.Words.Paragraph>()
    .FirstOrDefault(p => p.GetText().Trim() == "My Heading");

if (headingParagraph != null)
{
    // Position the DocumentBuilder at the heading paragraph
    var builder = new Aspose.Words.DocumentBuilder(doc);
    builder.MoveTo(headingParagraph);

    // Insert a new OLE object (e.g., an Excel file) after the heading.
    // The object will be displayed as an icon with a custom caption.
    // Use the overload that inserts from a file, specifies it as embedded (isLinked = false),
    // displays it as an icon (asIcon = true), and provides a custom icon image.
    using (var iconStream = System.IO.File.OpenRead(@"Images\CustomIcon.ico"))
    {
        builder.InsertOleObject(
            @"Data\SampleSpreadsheet.xlsx", // path to the OLE source file
            false,                         // isLinked = false (embed the file)
            true,                          // asIcon = true (show as icon)
            iconStream);                   // custom icon image stream
    }

    // Optionally, modify properties of the newly inserted OLE object.
    // The newly inserted shape is the last child of the document body.
    var oleShape = (Aspose.Words.Drawing.Shape)doc.GetChildNodes(Aspose.Words.NodeType.Shape, true)
        .Cast<Aspose.Words.Drawing.Shape>()
        .LastOrDefault(s => s.OleFormat != null);

    if (oleShape != null)
    {
        // Example: change the ProgId (if needed)
        oleShape.OleFormat.ProgId = "Excel.Sheet.12";

        // Example: lock the OLE object from automatic updates
        oleShape.OleFormat.IsLocked = true;
    }
}

// Save the modified document
doc.Save("Output.docx");
