using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first-page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Insert the logo image (provide the correct path to your image file).
        Shape logo = builder.InsertImage("logo.png");

        // Configure the image as a floating shape behind the text.
        logo.WrapType = WrapType.None;
        logo.BehindText = true;

        // Position the image relative to the page.
        logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        logo.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Center the image horizontally and vertically on the page.
        logo.HorizontalAlignment = HorizontalAlignment.Center;
        logo.VerticalAlignment = VerticalAlignment.Center;

        // Return the cursor to the main document body.
        builder.MoveToSection(0);

        // Add sample content to demonstrate the header.
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");

        // Save the document.
        doc.Save("FirstPageHeader.docx");
    }
}
