using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.EPPlus.Builder.Excel.Elements
{
    [TestClass]
    public class ParagraphFixture
    {

        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private ExcelParagraph _SystemUnderTest;
        public ExcelParagraph SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new ExcelParagraph();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void Paragraph_File_Basic()
        {
            // arrange
            var expected = true;

            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            var docMeta = new DocumentMetaData("ParagraphBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Paragraph Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.ParagraphBasic(docType, docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Paragraph_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Excel;
            IDocumentDirector director = new ContentReportBuilder.EPPlus.DocumentDirector();
            var docMeta = new DocumentMetaData("ParagraphBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Paragraph Basic Section";
            var docModel = TestData.ParagraphBasic(docType, docMeta, sectionTitle);
            string errorMsg;
            // act
            var mem = director.BuildStream(docType, docModel, out errorMsg);

            // assert
            Assert.AreNotEqual(mem, null);
            Assert.AreEqual("", errorMsg);
            StreamToFile.WriteStream(docType, docModel, mem);
        }
    }
}
