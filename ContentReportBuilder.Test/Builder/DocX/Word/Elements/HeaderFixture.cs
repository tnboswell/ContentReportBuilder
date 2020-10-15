using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.ElementModels;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using iTextSharp.text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    [TestClass]
    public class HeaderFixture
    {
        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private WordHeader _SystemUnderTest;
        public WordHeader SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new WordHeader();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void Header_Paragraph_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("HeaderParagraphBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Header Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.HeaderParagraphBasic(docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Header_Image_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("HeaderImageBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Header Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.HeaderImageBasic(docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Header_Paragraph_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("HeaderParagraphBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Header Basic Section";
            var docModel = TestData.HeaderParagraphBasic(docMeta, sectionTitle);
            string errorMsg;
            // act
            var mem = director.BuildStream(docType, docModel, out errorMsg);

            // assert
            Assert.AreNotEqual(mem, null);
            Assert.AreEqual("", errorMsg);

            StreamToFile.WriteStream(docType, docModel, mem);
        }

        [TestMethod]
        public void Header_Image_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("HeaderImageBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Header Basic Section";
            var docModel = TestData.HeaderImageBasic(docMeta, sectionTitle);
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
