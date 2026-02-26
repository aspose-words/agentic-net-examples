using System;
using Aspose.Words;
using Aspose.Words.Loading;   // RtfLoadOptions lives here
using Aspose.Words.Saving;    // RtfSaveOptions lives here

namespace LoadRtfAndProtectEquations
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an RTF document using RtfLoadOptions
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document doc = new Document(@"C:\Input\sample.rtf", loadOptions);

            // Apply read‑only protection to the whole document.
            // Aspose.Words does not expose a per‑OfficeMath read‑only flag, so protecting the
            // entire document ensures that equations cannot be edited in Microsoft Word.
            doc.Protect(ProtectionType.ReadOnly);

            // Save the protected document (keeping the RTF format).
            RtfSaveOptions saveOptions = new RtfSaveOptions();
            doc.Save(@"C:\Output\sample_protected.rtf", saveOptions);
        }
    }
}
