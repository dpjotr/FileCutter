using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace FileCutter
{
    class ExcelCooperator
    {
        internal string xlSourceFileName ;
        internal string xlRefFilename ;
        internal string resultDestination;
        internal ExcelCooperator(string source, string reference, string destination)
        {
            xlSourceFileName = source;
            xlRefFilename = reference;
            resultDestination = destination;
        }

        internal List<IXLRow> getRows(string fileName, int startRow, int worksheetNumber = 1)
        {
            List<IXLRow> rows = new List<IXLRow>();

            var workbook = new XLWorkbook(fileName);
            var worksheet = workbook.Worksheet(worksheetNumber);
            //Need to launch it parallel
            for (int i=startRow; !worksheet.Cell(i,1).IsEmpty(); i++)
                rows.Add(worksheet.Row(i));
           
            return rows;
        }

        internal IXLRow firstRow(string fileName)
        {
            List<IXLRow> rows = new List<IXLRow>();
            var workbook = new XLWorkbook(fileName);          

            return workbook.Worksheet(1).Row(1);
        }



        internal List<IXLCell> getCells(List<IXLRow> sourceRows, int column)
        {

            var cells = sourceRows
                .AsParallel()
                .Where(x => (x.Cell(3).GetString() == "C001" || x.Cell(3).GetString() == "B031"))
                .Select(x => x.Cell(column))
                .ToList();

            return cells;
        }

        internal void createFile(List<IXLRow> rows, int number=1)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Лист1");

            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < 18; j++)
                {
                    worksheet.Cell(i + 2, j + 1).Value =
                        rows[i].Cell(j + 1).Value;

                    worksheet.Cell(i + 2, j + 1).DataType =
                        rows[i].Cell(j + 1).DataType;

                    
                }                    

            workbook.SaveAs(resultDestination+"\\file_"+number+".xlsx");            
        }

        internal List<List<IXLRow>> divideList(List<IXLRow> source, int amountOfFiles) 
        {
            int averageLinesInFile = source.Count / amountOfFiles + 1;
            var groupped = source.GroupBy(x => x.Cell(7).Value.ToString());
            List<List<IXLRow>> result = new List<List<IXLRow>>();
            result.Add(new List<IXLRow>());
            int part = 0;

            foreach (var x in groupped)
            {
                foreach(var y in x)
                {
                    result[part].Add(y);
                }
                if (result[part].Count>averageLinesInFile)
                {
                    result.Add(new List<IXLRow>());
                    part++;
                }
            }

            return result;
        }
    }
}
