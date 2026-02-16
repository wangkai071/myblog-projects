#region 程序集 EPPlus, Version=4.5.3.3, Culture=neutral, PublicKeyToken=ea159fdaa78159a1
// D:\sqlserver2025\SQLserver导出excel01\packages\EPPlus.4.5.3.3\lib\net40\EPPlus.dll
#endregion

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Packaging;
using System;
using System.IO;

namespace OfficeOpenXml
{
    //
    // 摘要:
    //     Represents an Excel 2007/2010 XLSX file package. This is the top-level object
    //     to access all parts of the document.
    //
    // 言论：
    //     FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
    //     if (newFile.Exists)
    //     {
    //     newFile.Delete();  // ensures we create a new workbook
    //     newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
    //     }
    //     using (ExcelPackage package = new ExcelPackage(newFile))
    //     {
    //     // add a new worksheet to the empty workbook
    //     ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
    //     //Add the headers
    //     worksheet.Cells[1, 1].Value = "ID";
    //     worksheet.Cells[1, 2].Value = "Product";
    //     worksheet.Cells[1, 3].Value = "Quantity";
    //     worksheet.Cells[1, 4].Value = "Price";
    //     worksheet.Cells[1, 5].Value = "Value";
    //     //Add some items...
    //     worksheet.Cells["A2"].Value = "12001";
    //     worksheet.Cells["B2"].Value = "Nails";
    //     worksheet.Cells["C2"].Value = 37;
    //     worksheet.Cells["D2"].Value = 3.99;
    //     worksheet.Cells["A3"].Value = "12002";
    //     worksheet.Cells["B3"].Value = "Hammer";
    //     worksheet.Cells["C3"].Value = 5;
    //     worksheet.Cells["D3"].Value = 12.10;
    //     worksheet.Cells["A4"].Value = "12003";
    //     worksheet.Cells["B4"].Value = "Saw";
    //     worksheet.Cells["C4"].Value = 12;
    //     worksheet.Cells["D4"].Value = 15.37;
    //     //Add a formula for the value-column
    //     worksheet.Cells["E2:E4"].Formula = "C2*D2";
    //     //Ok now format the values;
    //     using (var range = worksheet.Cells[1, 1, 1, 5])
    //     {
    //     range.Style.Font.Bold = true;
    //     range.Style.Fill.PatternType = ExcelFillStyle.Solid;
    //     range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
    //     range.Style.Font.Color.SetColor(Color.White);
    //     }
    //     worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
    //     worksheet.Cells["A5:E5"].Style.Font.Bold = true;
    //     worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2,3,4,3).Address);
    //     worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
    //     worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";
    //     //Create an autofilter for the range
    //     worksheet.Cells["A1:E4"].AutoFilter = true;
    //     worksheet.Cells["A1:E5"].AutoFitColumns(0);
    //     // lets set the header text
    //     worksheet.HeaderFooter.oddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\"
    //     Inventory";
    //     // add the page number to the footer plus the total number of pages
    //     worksheet.HeaderFooter.oddFooter.RightAlignedText =
    //     string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
    //     // add the sheet name to the footer
    //     worksheet.HeaderFooter.oddFooter.CenteredText = ExcelHeaderFooter.SheetName;
    //     // add the file path to the footer
    //     worksheet.HeaderFooter.oddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath
    //     + ExcelHeaderFooter.FileName;
    //     worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
    //     worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];
    //     // Change the sheet view to show it in page layout mode
    //     worksheet.View.PageLayoutView = true;
    //     // set some document properties
    //     package.Workbook.Properties.Title = "Invertory";
    //     package.Workbook.Properties.Author = "Jan Källman";
    //     package.Workbook.Properties.Comments = "This sample demonstrates how to create
    //     an Excel 2007 workbook using EPPlus";
    //     // set some extended property values
    //     package.Workbook.Properties.Company = "AdventureWorks Inc.";
    //     // set some custom property values
    //     package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
    //     package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
    //     // save our new workbook and we are done!
    //     package.Save();
    //     }
    //     return newFile.FullName;
    //     More samples can be found at https://github.com/JanKallman/EPPlus/
    public sealed class ExcelPackage : IDisposable
    {
        //
        // 摘要:
        //     Maximum number of columns in a worksheet (16384).
        public const int MaxColumns = 16384;
        //
        // 摘要:
        //     Maximum number of rows in a worksheet (1048576).
        public const int MaxRows = 1048576;

        //
        // 摘要:
        //     Create a new instance of the ExcelPackage. Output is accessed through the Stream
        //     property, using the OfficeOpenXml.ExcelPackage.SaveAs(System.IO.FileInfo) method
        //     or later set the OfficeOpenXml.ExcelPackage.File property.
        public ExcelPackage();
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a stream
        //
        // 参数:
        //   newStream:
        //     The stream object can be empty or contain a package. The stream must be Read/Write
        public ExcelPackage(Stream newStream);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing file or creates
        //     a new file.
        //
        // 参数:
        //   newFile:
        //     If newFile exists, it is opened. Otherwise it is created from scratch.
        public ExcelPackage(FileInfo newFile);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing file or creates
        //     a new file.
        //
        // 参数:
        //   newFile:
        //     If newFile exists, it is opened. Otherwise it is created from scratch.
        //
        //   password:
        //     Password for an encrypted package
        public ExcelPackage(FileInfo newFile, string password);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing template.
        //     If newFile exists, it will be overwritten when the Save method is called
        //
        // 参数:
        //   newFile:
        //     The name of the Excel file to be created
        //
        //   template:
        //     The name of the Excel template to use as the basis of the new Excel file
        public ExcelPackage(FileInfo newFile, FileInfo template);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing template.
        //
        // 参数:
        //   template:
        //     The name of the Excel template to use as the basis of the new Excel file
        //
        //   useStream:
        //     if true use a stream. If false create a file in the temp dir with a random name
        public ExcelPackage(FileInfo template, bool useStream);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a stream
        //
        // 参数:
        //   newStream:
        //     The stream object can be empty or contain a package. The stream must be Read/Write
        //
        //   Password:
        //     The password to decrypt the document
        public ExcelPackage(Stream newStream, string Password);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a stream
        //
        // 参数:
        //   newStream:
        //     The output stream. Must be an empty read/write stream.
        //
        //   templateStream:
        //     This stream is copied to the output stream at load
        public ExcelPackage(Stream newStream, Stream templateStream);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing template.
        //     If newFile exists, it will be overwritten when the Save method is called
        //
        // 参数:
        //   newFile:
        //     The name of the Excel file to be created
        //
        //   template:
        //     The name of the Excel template to use as the basis of the new Excel file
        //
        //   password:
        //     Password to decrypted the template
        public ExcelPackage(FileInfo newFile, FileInfo template, string password);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a existing template.
        //
        // 参数:
        //   template:
        //     The name of the Excel template to use as the basis of the new Excel file
        //
        //   useStream:
        //     if true use a stream. If false create a file in the temp dir with a random name
        //
        //   password:
        //     Password to decrypted the template
        public ExcelPackage(FileInfo template, bool useStream, string password);
        //
        // 摘要:
        //     Create a new instance of the ExcelPackage class based on a stream
        //
        // 参数:
        //   newStream:
        //     The output stream. Must be an empty read/write stream.
        //
        //   templateStream:
        //     This stream is copied to the output stream at load
        //
        //   Password:
        //     Password to decrypted the template
        public ExcelPackage(Stream newStream, Stream templateStream, string Password);

        //
        // 摘要:
        //     Information how and if the package is encrypted
        public ExcelEncryption Encryption { get; }
        //
        // 摘要:
        //     The output stream. This stream is the not the encrypted package. To get the encrypted
        //     package use the SaveAs(stream) method.
        public Stream Stream { get; }
        //
        // 摘要:
        //     The output file. Null if no file is used
        public FileInfo File { get; set; }
        //
        // 摘要:
        //     Automaticlly adjust drawing size when column width/row height are adjusted, depending
        //     on the drawings editBy property. Default True
        public bool DoAdjustDrawings { get; set; }
        //
        // 摘要:
        //     Returns a reference to the workbook component within the package. All worksheets
        //     and cells can be accessed through the workbook.
        public ExcelWorkbook Workbook { get; }
        //
        // 摘要:
        //     Returns a reference to the package
        public ZipPackage Package { get; }
        //
        // 摘要:
        //     Compatibility settings for older versions of EPPlus.
        public CompatibilitySettings Compatibility { get; }
        //
        // 摘要:
        //     Compression option for the package
        public CompressionLevel Compression { get; set; }

        //
        // 摘要:
        //     Closes the package.
        public void Dispose();
        //
        // 摘要:
        //     Saves and returns the Excel files as a bytearray. Note that the package is closed
        //     upon save
        public byte[] GetAsByteArray();
        //
        // 摘要:
        //     Saves and returns the Excel files as a bytearray Note that the package is closed
        //     upon save
        //
        // 参数:
        //   password:
        //     The password to encrypt the workbook with. This parameter overrides the Encryption.Password.
        public byte[] GetAsByteArray(string password);
        //
        // 摘要:
        //     Loads the specified package data from a stream.
        //
        // 参数:
        //   input:
        //     The input.
        public void Load(Stream input);
        //
        // 摘要:
        //     Loads the specified package data from a stream.
        //
        // 参数:
        //   input:
        //     The input.
        //
        //   Password:
        //     The password to decrypt the document
        public void Load(Stream input, string Password);
        //
        // 摘要:
        //     Saves all the components back into the package. This method recursively calls
        //     the Save method on all sub-components. We close the package after the save is
        //     done.
        public void Save();
        //
        // 摘要:
        //     Saves all the components back into the package. This method recursively calls
        //     the Save method on all sub-components. The package is closed after it has ben
        //     saved d to encrypt the workbook with.
        //
        // 参数:
        //   password:
        //     This parameter overrides the Workbook.Encryption.Password.
        public void Save(string password);
        //
        // 摘要:
        //     Copies the Package to the Outstream The package is closed after it has been saved
        //
        // 参数:
        //   OutputStream:
        //     The stream to copy the package to
        public void SaveAs(Stream OutputStream);
        //
        // 摘要:
        //     Saves the workbook to a new file The package is closed after it has been saved
        //
        // 参数:
        //   file:
        //     The file
        //
        //   password:
        //     The password to encrypt the workbook with. This parameter overrides the Encryption.Password.
        public void SaveAs(FileInfo file, string password);
        //
        // 摘要:
        //     Saves the workbook to a new file The package is closed after it has been saved
        //
        // 参数:
        //   file:
        //     The file location
        public void SaveAs(FileInfo file);
        //
        // 摘要:
        //     Copies the Package to the Outstream The package is closed after it has been saved
        //
        // 参数:
        //   OutputStream:
        //     The stream to copy the package to
        //
        //   password:
        //     The password to encrypt the workbook with. This parameter overrides the Encryption.Password.
        public void SaveAs(Stream OutputStream, string password);
    }
}