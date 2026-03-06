using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Configure HTML save options to use an external CSS file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = "custom.css"
        };

        // Attach a callback that writes custom CSS content to the stylesheet file.
        saveOptions.CssSavingCallback = new CustomCssSavingCallback("custom.css",
            "/* Custom CSS */\n" +
            "body { font-family: Arial, sans-serif; margin: 20px; }\n" +
            "h1 { color: #2E8B57; }\n");

        // Save the document as HTML using the configured options.
        doc.Save("output.html", saveOptions);
    }
}

// Implements the CSS saving callback to provide custom stylesheet content.
public class CustomCssSavingCallback : ICssSavingCallback
{
    private readonly string _cssFilePath;
    private readonly string _cssContent;

    public CustomCssSavingCallback(string cssFilePath, string cssContent)
    {
        _cssFilePath = cssFilePath;
        _cssContent = cssContent;
    }

    public void CssSaving(CssSavingArgs args)
    {
        // Create or overwrite the CSS file.
        args.CssStream = new FileStream(_cssFilePath, FileMode.Create, FileAccess.Write);

        // Write the custom CSS text.
        using (StreamWriter writer = new StreamWriter(args.CssStream, Encoding.UTF8, 1024, leaveOpen: true))
        {
            writer.Write(_cssContent);
        }

        // Ensure the CSS is exported and the stream is closed after saving.
        args.IsExportNeeded = true;
        args.KeepCssStreamOpen = false;
    }
}
