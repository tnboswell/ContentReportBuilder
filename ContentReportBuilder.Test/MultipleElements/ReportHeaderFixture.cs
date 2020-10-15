using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.ElementModels;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Models.Styles;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.MultipleElements
{
    [TestClass]
    public class ReportHeaderFixture
    {

        [TestInitialize]
        public void OnTestInitialize()
        {
            _DocModel = null;
        }

        private DocumentModel _DocModel;
        public DocumentModel DocModel
        {
            get
            {
                if (_DocModel == null)
                {
                    string title = "Report Header";
                    _DocModel = new DocumentModel(new DocumentStyle(), new DocumentMetaData(TestConstants.OUTPUT_FILE, title, new List<string> { "Joe Bloggs" }));

                    var sectionMeta = new SectionMetaData(title);
                    var elementList = new List<DocumentModelElement>();
                    int colWidth = 8;
                    int headingWidth = (colWidth / 2);
                    int textWidth = (colWidth / 2) - 1;
                    int textStartRow = 1;
                    int nextRow = 1;
                    int continueRow = 0;
                    int startRow = 0;

                    elementList.Add(new ImageElementModel(new LocalImageData("testImage.jpg"), null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(3, 0), nextRow, startRow)
                    }));
                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Title" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Centre,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new FontStyle(),
                            Cell = new CellStyle(new CellMergingStyle(6, 0), continueRow, textStartRow)
                        }));

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Client" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));

                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "Some Client Name" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Task" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "Some Task Name" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Date" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "31 Jul 20" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Invoice Number" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "1234-P" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));
                    /*
                     *            |
                     *      Logo  |         Report Title
                     *            |
                     *      -------
                     *      Client        |
                     *      Task          |
                     *      Date          | 
                     *      Invoice Number|
                     */
                    var section = new DocumentModelSection(sectionMeta);
                    section.AddElementRange(elementList);
                    _DocModel.AddDocumentSection(section);
                }

                return _DocModel;
            }
        }

        [TestMethod]
        public void ReportHeader_File()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            DocModel.MetaData.Title += "_File";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, DocModel, out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void ReportHeader_Stream()
        {
            // arrange
            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            DocModel.MetaData.Title += "_Stream";
            string errorMsg;
            // act
            var mem = director.BuildStream(docType, DocModel, out errorMsg);

            // assert
            Assert.AreNotEqual(mem, null);
            Assert.AreEqual("", errorMsg);

            StreamToFile.WriteStream(docType, DocModel, mem);

        }
    }
}
