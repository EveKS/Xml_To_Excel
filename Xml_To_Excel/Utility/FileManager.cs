﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xml_To_Excel.Utility
{
    public interface IFileManager
    {
        bool IsExist(string filePath);
        bool IsFolderExist(string filePath);

        Task GetExcel(string filePath, string xmlsPath);
        Task MakeExcel(string excelPath, string xmlsPath, Encoding encoding);
    }

    public class FileManager : IFileManager
    {
        private readonly IExcelMaker _excelMaker;
        private readonly IReadXmlFoder _readXmlFolder;
        private readonly IExcelManager _excelManager;

        public FileManager() : this(new ExcelMaker(), new ReadXmlFoder(), new ExcelManager())
        {        }

        FileManager(IExcelMaker excelMaker, IReadXmlFoder readXmlFolder, IExcelManager excelManager)
        {
            _excelMaker = excelMaker;
            _readXmlFolder = readXmlFolder;
            _excelManager = excelManager;
        }
        private readonly Encoding _dafaultEncoding = Encoding.UTF8;

        public bool IsExist(string filePath)
        {
            bool isExist = File.Exists(filePath);
            return isExist;
        }
        public bool IsFolderExist(string folderPath)
        {
            bool isExist = Directory.Exists(folderPath);
            return isExist;
        }

        public async Task GetExcel(string filePath, string xmlsPath)
        {
            await MakeExcel(filePath, xmlsPath, _dafaultEncoding);
        }

        // Do Make
        public async Task MakeExcel(string excelPath, string xmlsPath, Encoding encoding)
            => await Task.Run(() =>
        {
            ExcelSelect ExcelSelect = new ExcelSelect();
            ExcelSelect.PathExelSelect = excelPath;
            ExcelSelect.SelectInExelFrom = "A1";
            ExcelSelect.SelectInExelTo = "Z500";
            var a = _excelManager.ListExcelArrayMaker(ExcelSelect);
            var b = _readXmlFolder.Read(xmlsPath, encoding);
            _excelMaker.ToMakeExcel(b, a);
        });
    }
}
