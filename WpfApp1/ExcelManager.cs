using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSampleApp
{
    public class ExcelManager
    {

        #region Data
        private SLDocument slDocExcelA;
        private SLDocument slDocExcelB;
        private SLDocument slDocExcelC;
        #endregion Data

        #region Loading Methods

        public List<ExcelA> LoadCollectionExcelA(String month)
        {
            try
            {
                slDocExcelA = new SLDocument("ExcelA.xlsx");
            }
            catch (System.IO.FileNotFoundException e)
            {
                slDocExcelA = new SLDocument();
                slDocExcelA.SaveAs("ExcelA.xlsx");
                slDocExcelA = new SLDocument("ExcelA.xlsx");
            }
            catch (System.IO.IOException f)
            {
                throw f;
            }

            _InitWorkSheets(slDocExcelA);
            slDocExcelA.SelectWorksheet(month);

            List<ExcelA> tempList = new List<ExcelA>();

            for (int i = 2; i < 80; i++)
            {
                ExcelA temp = new ExcelA();
                temp.Val1 = slDocExcelA.GetCellValueAsString("B" + i.ToString());
                temp.Val2 = slDocExcelA.GetCellValueAsString("C" + i.ToString());
                temp.Val3 = slDocExcelA.GetCellValueAsString("D" + i.ToString());

                if (string.IsNullOrEmpty(temp.Val1)) // Check for ID is not empty
                {
                    break;
                }
                else
                {
                    tempList.Add(temp);
                }
            }

            return tempList;
        }

        public List<ExcelB> LoadCollectionExcelB()
        {
            try
            {
                slDocExcelB = new SLDocument("ExcelB.xlsx");
            }
            catch (System.IO.FileNotFoundException e)
            {
                slDocExcelB = new SLDocument();
                slDocExcelB.SaveAs("ExcelB.xlsx");
                slDocExcelB = new SLDocument("ExcelB.xlsx");
            }
            catch (System.IO.IOException f)
            {
                throw f;
            }

            List<ExcelB> tempList = new List<ExcelB>();

            for (int i = 2; i < 80; i++)
            {
                ExcelB temp = new ExcelB();
                temp.Val3Main = slDocExcelB.GetCellValueAsString("B" + i.ToString());
                temp.Val4 = slDocExcelB.GetCellValueAsString("C" + i.ToString());
                temp.Val5 = slDocExcelB.GetCellValueAsString("D" + i.ToString());

                if (string.IsNullOrEmpty(temp.Val3Main)) // Check for ID is not empty
                {
                    break;
                }
                else
                {
                    tempList.Add(temp);
                }
            }

            return tempList;
        }

        public List<ExcelC> LoadCollectionExcelC(String month)
        {
            try
            {
                slDocExcelC = new SLDocument("ExcelC.xlsx");
            }
            catch (System.IO.FileNotFoundException e)
            {
                slDocExcelC = new SLDocument();
                slDocExcelC.SaveAs("ExcelC.xlsx");
                slDocExcelC = new SLDocument("ExcelC.xlsx");
            }
            catch (System.IO.IOException f)
            {
                throw f;
            }

            _InitWorkSheets(slDocExcelC);
            slDocExcelC.SelectWorksheet(month);

            List<ExcelC> tempList = new List<ExcelC>();

            for (int i = 2; i < 80; i++)
            {
                ExcelC temp = new ExcelC();
                temp.Val1 = slDocExcelC.GetCellValueAsString("B" + i.ToString());
                temp.Val2 = slDocExcelC.GetCellValueAsString("C" + i.ToString());
                temp.Val3 = slDocExcelC.GetCellValueAsString("D" + i.ToString());
                temp.Val4 = slDocExcelC.GetCellValueAsString("E" + i.ToString());

                if (string.IsNullOrEmpty(temp.Val1)) // Check for ID is not empty
                {
                    break;
                }
                else
                {
                    tempList.Add(temp);
                }
            }

            return tempList;
        }

        public List<String> LoadCollectionBMains(List<ExcelB> CollectionExcelB)
        {
            var temp = new List<String>();

            foreach (var x in CollectionExcelB)
            {
                temp.Add(x.Val3Main);
            }

            return temp;
        }

        public List<ExcelA> LoadCollectionForExcelBValue(String ExcelBVal, List<ExcelA> ExcelAList)
        {
            List<ExcelA> temp = new List<ExcelA>();

            foreach (var x in ExcelAList)
            {
                if (x.Val3 != null && x.Val3.Equals(ExcelBVal))
                    temp.Add(x);
            }

            return temp;
        }

        #endregion Loading Methods

        #region Storing Methods

        public void StoreCollectionA(List<ExcelA> collection, String month)
        {
            int i = 2, j = 1;

            slDocExcelA.SelectWorksheet(month);

            for (int x = 0; x < 50; x++)
            {
                if (x < collection.Count)
                {
                    slDocExcelA.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelA.SetCellValue("B" + i.ToString(), collection.ElementAt(x).Val1);
                    slDocExcelA.SetCellValue("C" + i.ToString(), collection.ElementAt(x).Val2);
                    slDocExcelA.SetCellValue("D" + i.ToString(), collection.ElementAt(x).Val3);
                }
                else
                {
                    slDocExcelA.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelA.SetCellValue("B" + i.ToString(), "");
                    slDocExcelA.SetCellValue("C" + i.ToString(), "");
                    slDocExcelA.SetCellValue("D" + i.ToString(), "");
                }

                i++;
                j++;
            }
            slDocExcelA.Save();
            slDocExcelA = new SLDocument("ExcelA.xlsx");
        }

        public void StoreCollectionB(List<ExcelB> collection)
        {
            int i = 2, j = 1;
            for (int x = 0; x < 50; x++)
            {
                if (x < collection.Count)
                {
                    slDocExcelB.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelB.SetCellValue("B" + i.ToString(), collection.ElementAt(x).Val3Main);
                    slDocExcelB.SetCellValue("C" + i.ToString(), collection.ElementAt(x).Val4);
                    slDocExcelB.SetCellValue("D" + i.ToString(), collection.ElementAt(x).Val5);
                }
                else
                {
                    slDocExcelB.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelB.SetCellValue("B" + i.ToString(), "");
                    slDocExcelB.SetCellValue("C" + i.ToString(), "");
                    slDocExcelB.SetCellValue("D" + i.ToString(), "");
                }

                i++;
                j++;
            }
            slDocExcelB.Save();
            slDocExcelB = new SLDocument("ExcelB.xlsx");
        }

        public void StoreCollectionC(List<ExcelC> collection, String month)
        {
            int i = 2, j = 1;

            slDocExcelC.SelectWorksheet(month);

            for (int x = 0; x < 50; x++)
            {
                if (x < collection.Count)
                {
                    slDocExcelC.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelC.SetCellValue("B" + i.ToString(), collection.ElementAt(x).Val1);
                    slDocExcelC.SetCellValue("C" + i.ToString(), collection.ElementAt(x).Val2);
                    slDocExcelC.SetCellValue("D" + i.ToString(), collection.ElementAt(x).Val3);
                    slDocExcelC.SetCellValue("E" + i.ToString(), collection.ElementAt(x).Val4);
                }
                else
                {
                    slDocExcelC.SetCellValue("A" + i.ToString(), j.ToString());
                    slDocExcelC.SetCellValue("B" + i.ToString(), "");
                    slDocExcelC.SetCellValue("C" + i.ToString(), "");
                    slDocExcelC.SetCellValue("D" + i.ToString(), "");
                    slDocExcelC.SetCellValue("E" + i.ToString(), "");
                }

                i++;
                j++;
            }
            slDocExcelC.Save();
            slDocExcelC = new SLDocument("ExcelC.xlsx");         
        }

        #endregion Storing Methods

        #region Opening Methods

        public void OpenExcelA(String month)
        {
            System.Diagnostics.Process.Start("ExcelA.xlsx");
        }

        public void OpenExcelB()
        {
            System.Diagnostics.Process.Start("ExcelB.xlsx");
        }

        public void OpenExcelC(String month)
        {
            System.Diagnostics.Process.Start("ExcelC.xlsx");
        }

        #endregion Opening Methods

        #region Private Methods

        private void _InitWorkSheets(SLDocument slDoc)
        {
            slDoc.DeleteWorksheet("Sheet1"); //delete default worksheet created by SpreadSheetLite-API
            foreach (var item in Constants.Months)
            {
                if (!slDoc.SelectWorksheet(item))
                    slDoc.AddWorksheet(item);
            }
        }

        #endregion Private Methods

      
          //Setting Font:

          //  SLStyle style = slDocExcelA.CreateStyle();
          //  style.SetFont("FontName", fontSize);
          //  style.Font.Underline = DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.Single;
          //  style.Font.Bold = true;
          //  slDocExcelA.SetRowStyle(RowIndex, style);
          //  slDocExcelA.SetCellStyle(stringReference, style);  

          // http://spreadsheetlight.com/sample-code/
    }
}
