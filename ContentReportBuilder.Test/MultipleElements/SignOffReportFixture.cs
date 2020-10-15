using ContentReportBuilder.Models;
using ContentReportBuilder.Models.ElementDataModels;
using ContentReportBuilder.Models.ElementModels;
using ContentReportBuilder.Models.MetaData;
using ContentReportBuilder.Models.Styles;
using ContentReportBuilder.Test.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ContentReportBuilder.MultipleElements
{
    [TestClass]
    public class SignoffReportFixture
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
                    string title = "Client Signoff Report";
                    string consultantName = "Joe Bloggs";
                    string taskName = "Cool Task";
                    DateTime startDate = DateTime.Now;
                    DateTime endDate = DateTime.Now;
                    string clientName = "Cool Client";
                    string imagePath = "testImage.jpg";
                    
                    _DocModel = new DocumentModel(new DocumentStyle(), new DocumentMetaData(TestConstants.OUTPUT_FILE, title, new List<string> { "Joe Bloggs" }));

                    var sectionMeta = new SectionMetaData(consultantName);
                    var elementList = new List<DocumentModelElement>();
                    int nextRow = 1;
                    int continueRow = 0;
                    int startRow = 0;
                    int headingMerge = 1;
                    int textMerge = 1;
                    var section = new DocumentModelSection(sectionMeta);
                    section.AddElement(new ImageElementModel(new LocalImageData(imagePath), null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(null, 2, startRow)
                    }));

                    section.AddElement(new HeadingElementModel(new HeadingData { Text = "Client Signoff Report" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Right,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new ContentReportBuilder.Models.Styles.FontStyle("Arial", 24),
                            Cell = new CellStyle(new CellMergingStyle(7, 3), continueRow, startRow)
                        }));

                    section.AddElement(new HeadingElementModel(new HeadingData { Text = "Start:" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Centre,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                            Cell = new CellStyle(new CellMergingStyle(headingMerge, 0), nextRow, startRow)
                        }));

                    section.AddElement(new ParagraphElementModel(new ParagraphData { Text = startDate.ToShortDateString() }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textMerge, 0), continueRow, 1)
                    }));

                    section.AddElement(new HeadingElementModel(new HeadingData { Text = "End:" }, null,
                    new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(headingMerge, 0), continueRow, 1)
                    }));

                    section.AddElement(new ParagraphElementModel(new ParagraphData { Text = endDate.ToShortDateString() }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(textMerge, 0), continueRow, 1)
                    }));

                    section.AddElement(new HeadingElementModel(new HeadingData { Text = "Consultant Name:" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Centre,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                            Cell = new CellStyle(new CellMergingStyle(headingMerge, 0), nextRow, startRow)
                        }));

                    section.AddElement(new ParagraphElementModel(new ParagraphData { Text = consultantName }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(5, 0), continueRow, 1)
                    }));

                    section.AddElement(new HeadingElementModel(new HeadingData { Text = "Task:" }, null,
                        new ExcelElementStyle
                        {
                            HorizontalAlignment = HorizontalAlignment.Centre,
                            VeritcalAlignment = VeritcalAlignment.Middle,
                            Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                            Cell = new CellStyle(new CellMergingStyle(headingMerge, 0), nextRow, startRow)
                        }));

                    section.AddElement(new ParagraphElementModel(new ParagraphData { Text = taskName }, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(5, 0), continueRow, 1)
                    }));
                    
                    var result = new List<SignoffReportModel> {
                        new SignoffReportModel {
                            Work = "Work",
                            ClientHours = 8,
                            Company = "Company",
                            Date = "Today"

                        }
                    };
      
                    var table = result.Cast<object>().ToList();

                    var colStyles = new Dictionary<int, IExcelElementStyle>
                    {
                        {1, new ExcelElementStyle
                            {
                                HorizontalAlignment = HorizontalAlignment.Centre,
                                VeritcalAlignment = VeritcalAlignment.Middle,
                                Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                                Cell = new CellStyle(new CellMergingStyle(2, 0),1,0)
                            } 
                        },
                        {2, new ExcelElementStyle
                            {
                                HorizontalAlignment = HorizontalAlignment.Centre,
                                VeritcalAlignment = VeritcalAlignment.Middle,
                                Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                                Cell = new CellStyle(new CellMergingStyle(1, 0),1,0)
                            }
                        },
                        {3, new ExcelElementStyle
                            {
                                HorizontalAlignment = HorizontalAlignment.Centre,
                                VeritcalAlignment = VeritcalAlignment.Middle,
                                Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                                Cell = new CellStyle(new CellMergingStyle(1, 0),1,0)
                            }
                        }
                    };

                    section.AddElement(new TableElementModel(new TableData
                    {
                        Rows = table
                    }, colStyles, null, new ExcelElementStyle
                    {
                        HorizontalAlignment = HorizontalAlignment.Centre,
                        VeritcalAlignment = VeritcalAlignment.Middle,
                        Font = new ContentReportBuilder.Models.Styles.FontStyle(),
                        Cell = new CellStyle(new CellMergingStyle(2, 0), 1, 0)
                    })
                    { HeadingColour = Color.Gray});

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
