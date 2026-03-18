using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertImageIntoFooter
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the image. The method returns the Shape that contains the image.
        Shape imageShape = builder.InsertImage(@"C:\Images\Logo.png");

        // Set the shape to be a floating picture (no text wrapping) so it can be positioned freely.
        imageShape.WrapType = WrapType.None;
        imageShape.BehindText = true;

        // Position the image relative to the footer margin.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Align the image to the left side of the footer.
        imageShape.Left = 0;
        imageShape.Top = 0;

        // Resize the shape to match the height of the footer.
        // The footer height can be approximated by the bottom margin of the page.
        // Here we use the page margin height as a proxy for the footer height.
        double footerHeight = builder.PageSetup.BottomMargin;
        imageShape.Height = footerHeight;
        // Preserve the aspect ratio by letting the width adjust automatically.
        imageShape.Width = -1; // -1 tells Aspose.Words to keep the original width proportionally.

        // Fit the image data to the shape frame so the picture scales to the shape size.
        imageShape.ImageData.FitImageToShape();

        // Save the document.
        doc.Save(@"C:\Output\DocumentWithFooterImage.docx");
    }
}
