using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class UpdateOleObjectExample
{
    static void Main()
    {
        // Load an existing DOCX document that contains an OLE object.
        Document doc = new Document("InputDocument.docx");

        // Find the first shape in the document that holds an OLE object.
        Shape oleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oleShape == null || oleShape.ShapeType != ShapeType.OleObject)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Access the OleFormat of the shape to modify its properties.
        OleFormat oleFormat = oleShape.OleFormat;

        // Example 1: Enable automatic update for linked OLE objects.
        // For embedded objects this property has no effect, but setting it is safe.
        oleFormat.AutoUpdate = true;

        // NOTE: IconCaption is a read‑only property in Aspose.Words, so it cannot be set directly.
        // If you need a different caption, you must replace the OLE object with a new one that has the desired caption.

        // Example 2: If the OLE object is an OLE Package (generic container), update its file name
        // and display name. This demonstrates how to modify the embedded data's metadata.
        if (oleFormat.OlePackage != null)
        {
            oleFormat.OlePackage.FileName = "UpdatedPackage.zip";
            oleFormat.OlePackage.DisplayName = "Updated Package.zip";
        }

        // Example 3: Replace the embedded data of the OLE object with new content.
        // We remove the old shape and insert a new OLE object (as an icon) using the file path.
        // The InsertOleObject overload that accepts a stream does not exist for the "as icon" variant,
        // therefore we use the file‑based overload.
        string newPackagePath = "NewPackage.zip"; // Ensure this file exists on disk.
        Node parent = oleShape.ParentNode;
        oleShape.Remove();

        DocumentBuilder builder = new DocumentBuilder(doc);
        // Move the builder to the position where the old shape was located.
        builder.MoveTo(parent);

        // Insert the new OLE object as an icon.
        Shape newOleShape = builder.InsertOleObjectAsIcon(
            newPackagePath,   // Path to the new package file.
            "Package",       // ProgId for a generic package.
            false,            // Not a linked object (embedded).
            null,             // Use the default icon.
            "New Package Icon"); // Custom icon caption.

        // Optionally set package properties for the new OLE object.
        if (newOleShape.OleFormat.OlePackage != null)
        {
            newOleShape.OleFormat.OlePackage.FileName = "NewPackage.zip";
            newOleShape.OleFormat.OlePackage.DisplayName = "NewPackage.zip";
        }

        // Save the modified document to a new file.
        doc.Save("UpdatedDocument.docx", SaveFormat.Docx);
    }
}
