using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using PdfSharpCore;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;

namespace FileAccess;
public class PDFLocator
{
    public void GetandProtectPDFFiles(string sourceDirectory)
    {

        try
        {
            var pdfFiles = Directory.EnumerateFiles(sourceDirectory, "*.pdf");
            //Check whether the pdf files belong to the employees in the summary list
            foreach(var pdfFile in pdfFiles)
            {
                using var document = PdfReader.Open(pdfFile);
                document.SecuritySettings.UserPassword = "12345";
                document.Save($"C:\\Test\\{new FileInfo(pdfFile).Name}");
            }
        }
        catch
        {

        }
    }
}
