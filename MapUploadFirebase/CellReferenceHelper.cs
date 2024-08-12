public static class CellReferenceHelper
{
    public static string GetCellReference(int row, int column)
    {
        string columnLetter = GetColumnLetter(column);
        return $"{columnLetter}{row}";
    }

    private static string GetColumnLetter(int column)
    {
        string columnLetter = string.Empty;
        while (column > 0)
        {
            int modulo = (column - 1) % 26;
            columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
            column = (column - 1) / 26;
        }
        return columnLetter;
    }
}

