using iTextSharp.text.pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using System;
using System.IO;
using System.Collections.Generic;

namespace OMNI.Helpers
{
    /// <summary>
    /// Map any electronic form to write to later.
    /// </summary>
    public class MapForm
    {
        /// <summary>
        /// Map a pdf from and save the results to a word document.
        /// </summary>
        public static void PdfForFields()
        {
            try
            {
                var fieldList = new List<string>();
                var ofd = new OpenFileDialog { Title = "Map pdf Form", Filter = "Adobe PDF|*.pdf", DefaultExt = "*.pdf", Multiselect = true };
                ofd.ShowDialog();
                foreach (string file in ofd.FileNames)
                {
                    fieldList.Clear();
                    using (PdfReader reader = new PdfReader(file))
                    {
                        var pdfField = reader.AcroFields;
                        foreach (var field in pdfField.Fields)
                        {
                            fieldList.Add($"Key = {field.Key}");
                            fieldList.Add($"Value = {pdfField.GetField(field.Key)}");
                        }
                    }
                    if (fieldList.Count > 0)
                    {
                        using (WordprocessingDocument _wpDoc = WordprocessingDocument.Create($"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\Mapped {Path.GetFileName(file)}.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                        {
                            var _wpMainDoc = _wpDoc.AddMainDocumentPart();
                            _wpMainDoc.Document = new Document();
                            var _wpDocBody = _wpMainDoc.Document.AppendChild(new Body());
                            var _wpDocPara = _wpDocBody.AppendChild(new Paragraph());
                            var _wpDocRun = _wpDocPara.AppendChild(new Run());
                            foreach (var s in fieldList)
                            {
                                _wpDocRun.AppendChild(new Text(s));
                                _wpDocRun.AppendChild(new Break());
                            }
                        }
                        System.Diagnostics.Process.Start($"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\Mapped {Path.GetFileName(file)}.docx");
                    }
                    else
                    {
                        ExceptionWindow.Show("Unable to Map", "The form you have selected either does not have mappable fields or the fields that are with in the form are corrupt and unreadable.\nPlease Contact IT if you feel you have reached this message in error.");
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
            }
        }

        public static void PdfForText()
        {
            try
            {
                var ofd = new OpenFileDialog { Title = "Map pdf Form", Filter = "Adobe PDF|*.pdf", DefaultExt = "*.pdf", Multiselect = true };
                ofd.ShowDialog();
                foreach (string file in ofd.FileNames)
                {
                    using (PdfReader reader = new PdfReader(file))
                    {
                        var test = reader.AcroForm;
                        File.WriteAllText("C:\\Users\\uif28100\\OneDrive - Continental AG\\Desktop\\testing.txt", reader.AcroForm.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionWindow.Show("Unhandled Exception", ex.Message, ex);
            }
        }
    }
}
