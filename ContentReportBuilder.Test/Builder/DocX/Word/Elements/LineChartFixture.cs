using ContentReportBuilder.Models;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ContentReportBuilder.docX.Builder.Word.Elements
{
    [TestClass]
    public class LineChartFixture
    {

        [TestInitialize]
        public void OnTestInitialize()
        {
            _SystemUnderTest = null;
        }

        private WordLineChart _SystemUnderTest;
        public WordLineChart SystemUnderTest
        {
            get
            {
                if (_SystemUnderTest == null)
                {
                    _SystemUnderTest = new WordLineChart();
                }

                return _SystemUnderTest;
            }
        }

        [TestMethod]
        public void LineChart_File_Basic()
        {
            // arrange
            var expected = true;
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("LineChartBasicFile", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Line Chart Basic Section";
            string errorMsg;
            // act
            var result = director.BuildFile(docType, TestData.LineChartBasic(docType, docMeta, sectionTitle), out errorMsg);

            // assert
            Assert.AreEqual(expected, result);
            Assert.AreEqual("", errorMsg);
        }

        [TestMethod]
        public void LineChart_Stream_Basic()
        {
            // arrange
            var docType = DocumentBuilderType.Word;
            IDocumentDirector director = new ContentReportBuilder.docX.DocumentDirector();
            var docMeta = new DocumentMetaData("LineChartBasicStream", new List<string> { "Troy Boswell" });
            docMeta.Path = TestConstants.OUTPUT_FILE;
            string sectionTitle = "Line Chart Basic Section";
            var docModel = TestData.BarChartBasic(docType, docMeta, sectionTitle);
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
