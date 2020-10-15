using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.ITextSharp.Builder.Pdf.Elements
{
    [TestClass]
    public class ImageFixture
    {
        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private PdfImage _SystemUnderTest;
        public PdfImage SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new PdfImage();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void Image_Url_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ImgUrlBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Image Url Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.ImgUrlBasic(docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Image_Url_Stream_Basic()
        {
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ImgUrlBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Image Url Basic Section";
            var docModel = TestData.ImgUrlBasic(docMeta, sectionTitle);
            string errorMsg;
            StreamToFile.WriteStream(docType, docModel, director.BuildStream(docType, docModel, out errorMsg));
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Image_Local_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ImgLocalBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Image Local Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.ImgLocalBasic(docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void Image_Local_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Pdf;
            IDocumentDirector director = new ContentReportBuilder.ITextSharp.DocumentDirector();
            var docMeta = new DocumentMetaData("ImgLocalBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Image Local Basic Section";
            var docModel = TestData.ImgLocalBasic(docMeta, sectionTitle);
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
