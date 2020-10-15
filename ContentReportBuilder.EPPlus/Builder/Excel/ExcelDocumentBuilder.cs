using ContentReportBuilder.Builder;
using ContentReportBuilder.EPPlus.Builder.Excel.Sections;
using ContentReportBuilder.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ContentReportBuilder.EPPlus.Builder.Excel
{
    public class ExcelDocumentBuilder : DocumentBuilderBase<ExcelWorksheet, ExcelWorkbook, ExcelWorksheet>, IDocumentBuilder
    {
        public ExcelDocumentBuilder() : base(new WorksheetBuilder())
        {
        }

        public override DocumentBuilderType Type => DocumentBuilderType.Excel;

        /// <summary>
        /// The method used to setup the document and call the document section builder
        /// </summary>
        /// <param name="document"></param>
        /// <param name="package"></param>
        /// <returns></returns>
        public override ExcelWorksheet Build(DocumentModel document, ExcelWorkbook package)
        {
            if (document.MetaData.Author.Count > 0)
            {
                package.Properties.Author = document.MetaData.Author.First();
            }

            package.Properties.Title = document.MetaData.Title;
            package.Properties.Created = DateTime.Now;
            package.Properties.Company = document.MetaData.Company;

            foreach (var section in document.DocumentSections ?? new List<DocumentModelSection>()) { 
                ExcelWorksheet newSheet = package.Worksheets.Add(section.MetaData.Title);

                newSheet.PrinterSettings.FooterMargin = (decimal)document.Style.FooterHeight;
                newSheet.PrinterSettings.HeaderMargin = (decimal)document.Style.HeaderHeight;
                newSheet.PrinterSettings.TopMargin = (decimal)document.Style.Margins.Top;
                newSheet.PrinterSettings.BottomMargin = (decimal)document.Style.Margins.Bottom;
                newSheet.PrinterSettings.LeftMargin = (decimal)document.Style.Margins.Left;
                newSheet.PrinterSettings.RightMargin = (decimal)document.Style.Margins.Right;
                _sectionBuilder.Build(Type, section, newSheet);
            }
            return null;
        }

        /// <summary>
        /// Build a valid excel document and save to file specified by the document meta data's path, filename and .xlsx extension.
        /// path/filename.xlsx
        /// </summary>
        /// <param name="document"></param>
        public override void BuildFile(DocumentModel document)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                Build(document, excelPackage.Workbook);

                FileInfo fi = new FileInfo(Filename(document.MetaData));
                excelPackage.SaveAs(fi);

            }
        }

        /// <summary>
        /// Build a stream for the valid excel document as a stream. Will be used to primarily as a web response.
        /// </summary>
        /// <param name="document"></param>
        /// <returns>Stream</returns>
        public override Stream BuildStream(DocumentModel document)
        {
            Stream mem = new MemoryStream();
            using (ExcelPackage excelPackage = new ExcelPackage())
            {

                Build(document, excelPackage.Workbook);

                excelPackage.SaveAs(mem);
                return mem;
            }
        }
    }
}
