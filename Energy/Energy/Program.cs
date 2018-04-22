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
                Console.Write("Введите № искомого файла elgamaRequest: ");
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
            //ArrayList for sum of different electrik cells

            ArrayList outCell2 = new ArrayList();
            ArrayList outCell6 = new ArrayList();
            ArrayList outCell27 = new ArrayList();

            //Filling Cell2
            for (var i=0; i<=rangeCells-1;i++)
            {
                cell2.Add(cells[i]);
            }
            //Return sum items of array cell2
            outCell2=SummAllItemsOfArray(cell2);

            //Filling Cell6
            for (var i=rangeCells;i<=(rangeCells*2)-1;i++)
            {
                cell6.Add(cells[i]);
            }
            //Return sum items of array cell6
            outCell6 = SummAllItemsOfArray(cell6);

            //Filling Cell27
            for (var i=rangeCells*2;i<= (cells.Count)-1;i++)
            {
                cell27.Add(cells[i]);
            }
            //Return sum items of array cell2
            outCell27 = SummAllItemsOfArray(cell27);

            using (StreamWriter sw = new StreamWriter(@"E:\SorceTreeRepositiry\EnergyCount\" + fullNameOfFile + ".txt", false, System.Text.Encoding.Default))
            {
                WriteAllDataToTextFile(outCell2, sw);

                WriteAllDataToTextFile(outCell6, sw);
                WriteAllDataToTextFile(outCell27, sw);
            }
            //Function for writing to .txt file
            void WriteAllDataToTextFile(ArrayList someArray, StreamWriter sw)
            {
                for(var i=0; i<=(someArray.Count)-1;i++)
                {
                    sw.WriteLine(someArray[i]);
                }
                sw.WriteLine("\n");
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

            Console.WriteLine("Данные успешно записаны в .txt файл");
            //Delay
            Console.ReadKey();

            //Internal function for summing of Array items
            ArrayList SummAllItemsOfArray(ArrayList sendArray)
            {
                int el = 3;
                int elements = el;
                int num = 0;
                float sum = 0;
                ArrayList internalArray = new ArrayList();
                while (num <= (sendArray.Count) && elements <= sendArray.Count)
                {
                    if (num < elements)
                    {
                        sum += (float)sendArray[num];
                        num++;
                    }
                    else
                    {
                        internalArray.Add(sum);
                        elements += el;
                        sum = 0;
                    }
                }
                return internalArray;
            }
        }
    }
}
