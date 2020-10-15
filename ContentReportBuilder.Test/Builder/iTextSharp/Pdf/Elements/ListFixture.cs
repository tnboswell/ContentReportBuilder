using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    [TestClass]
    public class ListFixture
    {

        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private PdfList _SystemUnderTest;
        public PdfList SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new PdfList();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void List_Ordered_File_Basic()
        {
            // arrange
            var ordered = true;
            var expected = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ListOrderedBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "List Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.ListBasic(docType, docMeta, sectionTitle, ordered), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void List_Ordered_Stream_Basic()
        {
            // arrange
            var ordered = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ListOrderedBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "List Basic Section";
            var docModel = TestData.ListBasic(docType, docMeta, sectionTitle, ordered);
            string errorMsg;
            // act
            var mem = director.BuildStream(docType, docModel, out errorMsg);

            // assert
            Assert.AreNotEqual(mem, null);
            Assert.AreEqual("", errorMsg);

            StreamToFile.WriteStream(docType, docModel, mem);
        }

        [TestMethod]
        public void List_Unordered_File_Basic()
        {
            // arrange
            var ordered = false;
            var expected = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ListUnorderedBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "List Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.ListBasic(docType, docMeta, sectionTitle, ordered), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void List_Unordered_Stream_Basic()
        {
            // arrange
            var ordered = false;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ListUnorderedBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "List Basic Section";
            var docModel = TestData.ListBasic(docType, docMeta, sectionTitle, ordered);
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
