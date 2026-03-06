using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // HTML template containing a conditional block (example uses VML conditional comments)
        const string html = @"
<html>
    <!--[if gte vml 1]>
        <img src='image_vml.jpg' />
    <![endif]-->
    <!--[if !vml]>
        <img src='image_png.png' />
    <![endif]-->
</html>";

        // Load the HTML using HtmlLoadOptions with the default BlockImportMode (Merge)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.BlockImportMode = BlockImportMode.Merge; // default value

        // Convert the HTML string to a stream for loading
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            // Create the document from the HTML stream and load options
            Document doc = new Document(stream, loadOptions);

            // Save the resulting document using SaveOptions (default template not required here)
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
            doc.Save("Output.docx", saveOptions);
        }
    }
}
