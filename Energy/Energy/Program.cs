using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Collections;

namespace Energy
{
    class Program
    {
        public const string NameOfFile = "elgamaRequest";
        //number of downloaded file elgamaRequest
        internal static string number = null;

        static void Main(string[] args)
        {
            string fullNameOfFile = null;
           
            while(string.IsNullOrEmpty(number))
            {
                Console.WriteLine("Введите № искомого файла elgamaRequest\n");
                number = Console.ReadLine();
            }

            fullNameOfFile = NameOfFile + number;


            Application xl = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = xl.Workbooks.Open(@"E:\SorceTreeRepositiry\EnergyCount\" + fullNameOfFile);
            Worksheet sheet = workbook.Sheets[1];
            
            int numRows = sheet.UsedRange.Rows.Count; //Rows quantity
            int numColumns = 6;     // quantity of Columns

            string[] record=new string [numRows];
            ArrayList cells = new ArrayList();                     
            
            
            for (int rowIndex = 7; rowIndex <= numRows; rowIndex++)  
            {
                for (int colIndex = 6; colIndex <= numColumns; colIndex++)
                {
                    Range cell = (Range)sheet.Cells[rowIndex, colIndex];
                    if (cell.Value != null)
                    {
                        string prom=(Convert.ToString(cell.Value)).Replace(".",",");
                        float res;
                        bool isDigit = float.TryParse(prom, out res);
                            if (isDigit && res!=0)
                            {
                                cells.Add(res);                                
                            }                           
                    }                   
                }               
            }
            int rangeCells = (cells.Count)/3;
            //ArrayLists for different electrik cells
            ArrayList cell2 = new ArrayList();
            ArrayList cell6 = new ArrayList();
            ArrayList cell27 = new ArrayList();

            //Filling Cell2
            for (var i=0; i<=rangeCells-1;i++)
            {
                cell2.Add(cells[i]);
            }

            // outCell2 = new float[cell2.Count/3];
            ArrayList outCell2 = new ArrayList();
            int el = 3;
            int elements = el;
            int num = 0;
            float sum = 0;
            while(num <= (cell2.Count) && elements<=cell2.Count)
            {
                if (num < elements)
                {
                    sum += (float)cell2[num];
                    num++;
                }
                else
                {
                    outCell2.Add(sum);

                    elements += el;
                    //if (elements > cell2.Count)
                    //    elements = cell2.Count;
                    sum = 0;
                }
            }
            
            //float sum = 0;
            //for (var i=0; i<=cell2.Count;i++)
            //{
            //    if(i<elements)
            //    {
            //        //Get by 3 elements of array           
            //        sum += (float)cell2[i];
            //    }

            //}

            //Filling Cell6
            for (var i=rangeCells;i<=(rangeCells*2)-1;i++)
            {
                cell6.Add(cells[i]);
            }

            //Filling Cell27
            for (var i=rangeCells*2;i<= (cells.Count)-1;i++)
            {
                cell27.Add(cells[i]);
            }

            //Settings for finish work with Microsoft Excel
            #region
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            
            //release com objects to fully kill excel process from running in the background            
            Marshal.ReleaseComObject(sheet);

            //close and release
            workbook.Close();
            Marshal.ReleaseComObject(workbook);

            //quit and release
            xl.Quit();
            Marshal.ReleaseComObject(xl);
            #endregion

            foreach (var j in cells)
            {
                Console.WriteLine(j+"\n");
            }

            //Delay
            Console.ReadKey();
        }
    }
}
