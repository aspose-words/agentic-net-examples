using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class FitHeightImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // URL of the image to be inserted.
        string imageUrl = "https://example.com/sample-image.jpg";

        // Insert an INCLUDEPICTURE field. The field code contains the image URL
        // followed by the \\* MERGEFORMAT switch and the \\* MERGEFORMAT switch
        // with the fitHeight option (\\* MERGEFORMAT \\* MERGEFORMAT is the
        // representation of the fitHeight switch in the field code).
        // The field will be linked so that the image is not stored in the document.
        FieldIncludePicture includePicture = (FieldIncludePicture)builder.InsertField(FieldType.FieldIncludePicture, true);
        includePicture.SourceFullName = $"{imageUrl} \\* MERGEFORMAT \\* MERGEFORMAT";
        includePicture.IsLinked = true;

        // Enable vertical resizing so the picture fits the height of the paragraph.
        includePicture.ResizeVertically = true;
        // Optionally enable horizontal resizing as well.
        includePicture.ResizeHorizontally = true;

        // Update fields to fetch the image and apply the fitHeight behavior.
        doc.UpdateFields();

        // Save the document.
        doc.Save("FitHeightImage.docx");
    }
}
