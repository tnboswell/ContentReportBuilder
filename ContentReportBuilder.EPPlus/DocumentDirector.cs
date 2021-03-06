﻿using ContentReportBuilder.Builder;
using ContentReportBuilder.EPPlus.Builder.Excel;
using ContentReportBuilder.Models;
using System;
using System.IO;

namespace ContentReportBuilder.EPPlus
{
    public class DocumentDirector : DocumentDirectorBase, IDocumentDirector
    {  
        /// <summary>
        /// Gets a requested document builder type
        /// </summary>
        /// <param name="type"></param>
        /// <returns>returns the requested document builder</returns>
        protected override IDocumentBuilder GetBuilder(DocumentBuilderType type)
        {
            if (type == DocumentBuilderType.Excel)
            {
                return new ExcelDocumentBuilder();
            }

            return null;
        }
    }
}
