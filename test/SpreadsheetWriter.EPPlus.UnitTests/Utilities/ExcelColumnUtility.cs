﻿using System;

namespace SpreadsheetWriter.EPPlus.UnitTests.Utilities
{
    public static class ExcelColumnUtility
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = ((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
