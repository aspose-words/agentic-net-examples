using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world!");

        var options = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        doc.Save("Output.html", options);
    }
}
