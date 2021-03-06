using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    [TestClass]
    public class HeadingFixture
    {
        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private PdfHeading _SystemUnderTest;
        public PdfHeading SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new PdfHeading();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void Heading_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("HeadingBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Heading Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.HeadingBasic(docType, docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Heading_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("HeadingBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Heading Basic Section";
            var docModel = TestData.HeadingBasic(docType, docMeta, sectionTitle);
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
