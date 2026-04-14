using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create deterministic sample images and store them as files.
        const int imageCount = 2;
        string[] imageFiles = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string fileName = $"sample{i + 1}.png";
            CreateSampleImage(100, 100, fileName);
            imageFiles[i] = fileName;
        }

        // Step 2: Build a simple DOCX template containing image merge fields.
        const string templatePath = "template.docx";
        BuildTemplateDocument(templatePath);

        // Step 3: Simulate a database table that holds image BLOBs.
        DataTable imageTable = new DataTable();
        imageTable.Columns.Add("Image1", typeof(byte[]));
        imageTable.Columns.Add("Image2", typeof(byte[]));

        DataRow row = imageTable.NewRow();
        row["Image1"] = File.ReadAllBytes(imageFiles[0]);
        row["Image2"] = File.ReadAllBytes(imageFiles[1]);
        imageTable.Rows.Add(row);

        // Step 4: Load the template, configure the mail‑merge callback, and execute the merge.
        Document doc = new Document(templatePath);
        doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();
        doc.MailMerge.Execute(imageTable);

        // Step 5: Save the merged document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Step 6: Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file '{outputPath}'.");

        // Clean up temporary files (optional).
        foreach (string file in imageFiles)
            File.Delete(file);
        File.Delete(templatePath);
    }

    // Creates a deterministic PNG image using Aspose.Drawing and saves it to the specified path.
    private static void CreateSampleImage(int width, int height, string filePath)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Additional deterministic drawing can be added here if desired.
            bitmap.Save(filePath);
        }
    }

    // Generates a DOCX file that contains two image merge fields: Image1 and Image2.
    private static void BuildTemplateDocument(string filePath)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert first image merge field.
        builder.InsertField(" MERGEFIELD Image1 ");
        builder.Writeln();

        // Insert second image merge field.
        builder.InsertField(" MERGEFIELD Image2 ");

        template.Save(filePath);
    }

    // Implements the callback that supplies image streams to the mail‑merge engine.
    private class ImageFieldMergingHandler : IFieldMergingCallback
    {
        // Required by the interface but not used for text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

        // Called when a MERGEFIELD with an image tag is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The field value is expected to be a byte[] containing the image data.
            byte[] imageBytes = args.FieldValue as byte[];
            if (imageBytes == null)
                throw new InvalidOperationException("Expected image data as a byte array.");

            // Provide the image stream to the mail‑merge engine.
            args.ImageStream = new MemoryStream(imageBytes);
        }
    }
}
