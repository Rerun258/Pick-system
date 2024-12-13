using System;
using System.Formats.Asn1;
using FastExcel;
using System.IO;
using System.Linq;

// Get the input file path
var inputFile = new FileInfo("Order_details.xlsx");

// List of cell indices to print (adjust these indices as needed)
var selectedIndices = new[] {  2, 3,4, 5, 9};

// Create an instance of Fast Excel
using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
{
    foreach (var worksheet in fastExcel.Worksheets)
    {
        Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index));

        // To read the rows call read
        worksheet.Read();
        var rows = worksheet.Rows.ToList();

        // Print the content of selected cells in each row
        foreach (var row in rows)
        {
            var cells = row.Cells.ToArray();
            foreach (var index in selectedIndices)
            {
                if (index - 1 < cells.Length)
                {
                    var cell = cells[index - 1];
                    Console.Write(cell.Value + "\t");
                }
                else
                {
                    Console.Write("N/A\t"); // If the cell is not found at the given index
                }
            }
            Console.WriteLine();
        }

        Console.WriteLine(string.Format("Worksheet Rows:{0}", rows.Count()));
    }
}









//using System;
// using System.Formats.Asn1;
// using FastExcel;
// using System.IO;
// using System.Linq;

// // Get the input file path
// var inputFile = new FileInfo("Order_details.xlsx");
// var selectedIndices = new[] {0, 1, 2, 3, 10};

// // Create an instance of Fast Excel
// using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
// {
//     { foreach (var worksheet in fastExcel.Worksheets) { Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index)); // To read the rows call read worksheet.Read(); var rows = worksheet.Rows.ToList(); // Print the content of selected cells in each row foreach (var row in rows) { foreach (var index in selectedIndices) { var cell = row.Cells.FirstOrDefault(c => c.ColumnIndex == index); if (cell != null) { Console.Write(cell.Value + "\t"); } else { Console.Write("N/A\t"); // If the cell is not found at the given index } } Console.WriteLine(); } Console.WriteLine(string.Format("Worksheet Rows:{0}"
// }
// using System;
// using System.Formats.Asn1;
// using FastExcel;


// // Get the input file path
// var inputFile = new FileInfo("Order_details.xlsx");

// // Create an instance of Fast Excel
// using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
// {
//     foreach (var worksheet in fastExcel.Worksheets)
//     {
//         Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index));
        
//         //To read the rows call read
//         worksheet.Read();
//         var rows = worksheet.Rows.ToList();
//         //Do something with rows

//         Console.WriteLine(rows);
        
//         // for(int i=0;i<rows.count();i++)
//         // {
//         // Console.WriteLine(rows[i].ToString());
//         // }

//         foreach (var row in rows)
//         {
//             // Console.WriteLine(row.GetType);
//             Console.WriteLine(row.ToString());

//         }
//         Console.WriteLine(string.Format("Worksheet Rows:{0}", rows.Count()));
//     }
// }

// WorkBook workbook = WorkBook.Load("Order_details.xlsx");
// WorkSheet sheet = workbook.GetWorkSheet("Order Details - 2023-09-20T0822");
// foreach(var cell in sheet["A1 : A10"]){
//     Console.WriteLine(cell);

// }