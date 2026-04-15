using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;

namespace OleObjectSizeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare some dummy data to embed as an OLE package.
            // Here we use a simple text file content.
            byte[] dummyData = System.Text.Encoding.UTF8.GetBytes("Hello, Aspose.Words OLE object!");
            using (MemoryStream oleStream = new MemoryStream(dummyData))
            {
                // Insert the OLE object into the document.
                // Use the "Package" progId to embed generic data.
                // The object is inserted as content (asIcon = false) and without a custom presentation image.
                Shape oleShape = builder.InsertOleObject(oleStream, "Package", false, null);

                // Retrieve the display size of the OLE object (in points).
                double originalWidth = oleShape.Width;
                double originalHeight = oleShape.Height;

                Console.WriteLine($"Original OLE display size: Width = {originalWidth} pt, Height = {originalHeight} pt");

                // Adjust the display size – for example, increase both dimensions by 50%.
                oleShape.Width = originalWidth * 1.5;
                oleShape.Height = originalHeight * 1.5;

                Console.WriteLine($"Adjusted OLE display size: Width = {oleShape.Width} pt, Height = {oleShape.Height} pt");

                // If the OLE object is an ActiveX control, we can also adjust its internal size.
                // This part is optional and will be executed only when an OleControl is present.
                OleFormat oleFormat = oleShape.OleFormat;
                if (oleFormat?.OleControl != null)
                {
                    Forms2OleControl oleControl = (Forms2OleControl)oleFormat.OleControl;
                    // Retrieve current control size.
                    double controlOriginalWidth = oleControl.Width;
                    double controlOriginalHeight = oleControl.Height;

                    Console.WriteLine($"Original ActiveX control size: Width = {controlOriginalWidth} pt, Height = {controlOriginalHeight} pt");

                    // Increase control size by the same factor.
                    oleControl.Width = controlOriginalWidth * 1.5;
                    oleControl.Height = controlOriginalHeight * 1.5;

                    Console.WriteLine($"Adjusted ActiveX control size: Width = {oleControl.Width} pt, Height = {oleControl.Height} pt");
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OleObjectDemo.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
