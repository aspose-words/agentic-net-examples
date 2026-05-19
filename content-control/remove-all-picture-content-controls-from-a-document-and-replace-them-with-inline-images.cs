using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string imagePath = "sample.png";

        // Create a tiny PNG image to use in the example.
        CreateSamplePng(imagePath);

        // -----------------------------------------------------------------
        // Create a document that contains picture content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First picture content control.
        StructuredDocumentTag pictureSdt1 = builder.InsertStructuredDocumentTag(SdtType.Picture);
        builder.InsertImage(imagePath);
        builder.Writeln();

        // Second picture content control.
        StructuredDocumentTag pictureSdt2 = builder.InsertStructuredDocumentTag(SdtType.Picture);
        builder.InsertImage(imagePath);
        builder.Writeln();

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and replace picture content controls with inline images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Collect all picture SDTs.
        List<StructuredDocumentTag> pictureTags = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        // Remove each picture content control but keep its inner image.
        foreach (StructuredDocumentTag sdt in pictureTags)
        {
            sdt.RemoveSelfOnly();
        }

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }

    // Generates a 1x1 red PNG image from a Base64 string.
    private static void CreateSamplePng(string filePath)
    {
        const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcV8AAAAASUVORK5CYII=";
        byte[] bytes = Convert.FromBase64String(base64);
        File.WriteAllBytes(filePath, bytes);
    }
}
