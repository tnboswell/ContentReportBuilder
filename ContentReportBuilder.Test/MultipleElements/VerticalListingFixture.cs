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
    public class VerticalListingFixture
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
                    string title = "Vertical Listing";
                    _DocModel = new DocumentModel(new DocumentStyle(), new DocumentMetaData(TestConstants.OUTPUT_FILE, title, new List<string> { "Joe Bloggs" }));

                    var sectionMeta = new SectionMetaData(title);
                    var elementList = new List<DocumentModelElement>();
                    // report details
                    int colWidth = 8;
                    int headingWidth = (colWidth / 2);
                    int textWidth = (colWidth / 2) - 1;
                    int textStartRow = 1;
                    int nextRow = 1;
                    int continueRow = 0;
                    int startRow = 0;
                    /*
                     *              Title
                     *      Unit        |
                     *      Serial      |
                     *      Status      | 
                     *      Result      |
                     *      */

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Title" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Centre,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new FontStyle(),
                            Cell = new CellStyle(new CellMergingStyle(colWidth, 0), nextRow, startRow)
                        }));


                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Unit" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));

                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "Your Unit" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));

                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Serial" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "20-01-01" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));
                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Status" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "Complete" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));
                    elementList.Add(new HeadingElementModel(new HeadingData { Text = "Result" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingWidth, 0), nextRow, startRow)
                    }));
                    elementList.Add(new ParagraphElementModel(new ParagraphData { Text = "Passed" }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textWidth, 0), continueRow, textStartRow)
                    }));

                    var section = new DocumentModelSection(sectionMeta);
                    section.AddElementRange(elementList);
                    _DocModel.AddDocumentSection(section);
                }


                return _DocModel;
            }
        }
        [TestMethod]
        public void VertialListing_File()
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
        public void VertialListing_Stream()
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
