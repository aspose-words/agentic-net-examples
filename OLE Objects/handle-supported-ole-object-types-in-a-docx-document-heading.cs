using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

class OleObjectsDemo
{
    static void Main()
    {
        // Define directories (adjust as needed).
        string MyDir = @"C:\Data\";
        string ImageDir = Path.Combine(MyDir, "Images\\");
        string ArtifactsDir = Path.Combine(MyDir, "Output\\");
        Directory.CreateDirectory(ArtifactsDir);

        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 3. Insert a heading for the OLE objects section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Supported OLE Objects");

        // Reset style to normal for the following content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // -----------------------------------------------------------------
        // Insert an embedded Excel spreadsheet (from a file) – displayed as content.
        // -----------------------------------------------------------------
        string excelPath = Path.Combine(MyDir, "Spreadsheet.xlsx");
        builder.Writeln("Embedded Excel Spreadsheet:");
        // The file extension is used to detect the OLE type automatically.
        builder.InsertOleObject(excelPath, isLinked: false, asIcon: false, presentation: null);

        // -----------------------------------------------------------------
        // Insert an embedded PowerPoint presentation – displayed as an icon with a custom image.
        // -----------------------------------------------------------------
        string pptPath = Path.Combine(MyDir, "Presentation.pptx");
        string iconPath = Path.Combine(ImageDir, "Logo icon.ico");
        builder.Writeln("\nEmbedded PowerPoint Presentation (as icon):");
        builder.InsertOleObjectAsIcon(pptPath, isLinked: false, iconFile: iconPath, iconCaption: "My Presentation");

        // -----------------------------------------------------------------
        // Insert a generic file (ZIP archive) using the OLE Package mechanism – displayed as an icon.
        // -----------------------------------------------------------------
        string zipPath = Path.Combine(MyDir, "Archive.zip");
        builder.Writeln("\nEmbedded ZIP Archive (as OLE Package icon):");
        Shape zipShape = builder.InsertOleObjectAsIcon(zipPath, isLinked: false, iconFile: iconPath, iconCaption: "My Archive");

        // Set OLE Package properties for the ZIP file.
        zipShape.OleFormat.OlePackage.FileName = "Archive.zip";
        zipShape.OleFormat.OlePackage.DisplayName = "Archive.zip";

        // -----------------------------------------------------------------
        // Save the document containing the OLE objects.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(ArtifactsDir, "SupportedOleObjects.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the saved document and extract each embedded OLE object's raw data.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        // Get only Shape nodes (OLE objects are stored as Shape nodes).
        Shape[] oleShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                    .OfType<Shape>()
                                    .ToArray();

        foreach (Shape shape in oleShapes)
        {
            if (shape.OleFormat == null)
                continue; // Not an OLE object.

            OleFormat ole = shape.OleFormat;

            // Skip linked objects – they have no embedded data to extract.
            if (ole.IsLink)
                continue;

            // Determine a suitable file name using the suggested extension.
            string suggestedExt = ole.SuggestedExtension ?? ".bin";
            string outputFile = Path.Combine(ArtifactsDir,
                $"Extracted_{Guid.NewGuid()}{suggestedExt}");

            // Save the OLE data to the file system.
            ole.Save(outputFile);
        }
    }
}
