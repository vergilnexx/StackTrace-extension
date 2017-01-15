/***************************************************************************
 
Copyright (c) Microsoft Corporation. All rights reserved.
THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

***************************************************************************/

using System;
using System.Runtime.Serialization;

namespace Scanner
{
    [Serializable]
    class ScanHit : IScanHit, ISerializable
    {
        const string FilePathKey = "FilePath";
        const string LineKey = "Line";
        const string ColumnKey = "Column";
        const string TermKey = "Term";
        const string LineTextKey = "LineText";
        const string WarningKey = "Warning";

        readonly string _filePath;
        readonly int _line;
        readonly int _column;
        readonly string _searchingWord;
        readonly string _lineText;
        readonly string _warning;

        public ScanHit(string filePath, int line, int column, string lineText, string searchingWord, string warning)
        {
            _filePath = filePath;
            _line = line;
            _column = column;
            _searchingWord = searchingWord;
            _lineText = lineText;
            _warning = warning;
        }

        protected ScanHit(SerializationInfo info, StreamingContext context)
        {
            _filePath = info.GetString(FilePathKey);
            _line = info.GetInt32(LineKey);
            _column = info.GetInt32(ColumnKey);
            _searchingWord = (string)info.GetValue(TermKey, typeof(string));
            _lineText = info.GetString(LineTextKey);
            _warning = info.GetString(WarningKey);
        }

        #region IScanHit Members

        public string FilePath
        {
            get { return _filePath; }
        }

        public int Line
        {
            get { return _line; }
        }

        public int Column
        {
            get { return _column; }
        }

        public string LineText
        {
            get { return _lineText; }
        }

        public string SearchingWord
        {
            get { return _searchingWord; }
        }

        public string Warning
        {
            get { return _warning; }
        }

        #endregion IScanHit Members

        #region ISerializable Members

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue(FilePathKey, _filePath);
            info.AddValue(LineKey, _line);
            info.AddValue(ColumnKey, _column);
            info.AddValue(TermKey, _searchingWord);
            info.AddValue(LineTextKey, _lineText);
            info.AddValue(WarningKey, _warning);
        }

        #endregion ISerializable Members
    }
}
