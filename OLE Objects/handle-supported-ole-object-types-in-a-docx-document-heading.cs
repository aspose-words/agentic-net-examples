using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class OleObjectsDemo
{
    static void Main()
    {
        // Paths – adjust these to point to existing files on your machine.
        string dataDir = @"C:\Data\";
        string artifactsDir = @"C:\Artifacts\";
        string excelFile = Path.Combine(dataDir, "Sample.xlsx");
        string pptFile = Path.Combine(dataDir, "Sample.pptx");
        string zipFile = Path.Combine(dataDir, "Sample.zip");
        string iconFile = Path.Combine(dataDir, "Icon.ico");

        // -------------------------------------------------
        // 1. Create a new blank document and a DocumentBuilder.
        // -------------------------------------------------
        Document doc = new Document();                     // create
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 2. Insert an embedded Excel spreadsheet (normal view).
        // -------------------------------------------------
        builder.Writeln("Embedded Excel object:");
        using (FileStream excelStream = File.OpenRead(excelFile))
        {
            // InsertOleObject(stream, progId, asIcon, presentation)
            // progId "Excel.Sheet.12" corresponds to modern .xlsx files.
            builder.InsertOleObject(excelStream, "Excel.Sheet.12", false, null);
        }

        // -------------------------------------------------
        // 3. Insert a PowerPoint presentation as an icon with a custom caption.
        // -------------------------------------------------
        builder.Writeln("\nPowerPoint object as icon:");
        using (FileStream pptStream = File.OpenRead(pptFile))
        {
            // InsertOleObjectAsIcon(stream, progId, iconFile, iconCaption)
            builder.InsertOleObjectAsIcon(pptStream, "PowerPoint.Application", iconFile, "My Presentation");
        }

        // -------------------------------------------------
        // 4. Insert a generic file (ZIP) using the OLE Package mechanism.
        // -------------------------------------------------
        builder.Writeln("\nZIP archive as OLE Package:");
        byte[] zipBytes = File.ReadAllBytes(zipFile);
        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // progId "Package" tells Aspose.Words to treat the data as an OLE Package.
            Shape zipShape = builder.InsertOleObject(zipStream, "Package", false, null);
            // Set display name and file name for the package.
            zipShape.OleFormat.OlePackage.FileName = "Archive.zip";
            zipShape.OleFormat.OlePackage.DisplayName = "Sample Archive";
        }

        // -------------------------------------------------
        // 5. Save the document.
        // -------------------------------------------------
        string outDoc = Path.Combine(artifactsDir, "OleObjectsDemo.docx");
        doc.Save(outDoc);                                 // save

        // -------------------------------------------------
        // 6. Load the saved document (demonstrating LoadOptions usage).
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions { IgnoreOleData = false };
        Document loadedDoc = new Document(outDoc, loadOptions); // load

        // -------------------------------------------------
        // 7. Extract each embedded OLE object to a separate file.
        // -------------------------------------------------
        Shape[] oleShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                    .ToArray()
                                    .OfType<Shape>()
                                    .Where(s => s.OleFormat != null)
                                    .ToArray();

        foreach (Shape shape in oleShapes)
        {
            OleFormat ole = shape.OleFormat;

            // Determine a suitable file name.
            string suggestedExt = ole.SuggestedExtension; // e.g. ".xlsx", ".pptx", ".zip"
            string baseName = Path.GetFileNameWithoutExtension(ole.SuggestedFileName);
            if (string.IsNullOrEmpty(baseName))
                baseName = "ExtractedObject";

            string outPath = Path.Combine(artifactsDir, baseName + suggestedExt);

            // Save the OLE data to the file system.
            ole.Save(outPath); // save
        }

        Console.WriteLine("OLE objects inserted and extracted successfully.");
    }
}
