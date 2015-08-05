using System.Windows;
using MarkupConverter;
using Microsoft.Office.Interop.Excel;
using System; 

namespace MarkupConverter
{

    public partial class Window1 : System.Windows.Window
    {
        private IMarkupConverter markupConverter;
        public Window1()
        {
            markupConverter = new MarkupConverter();
        }

        public void convertHTMLToXAML(object sender, RoutedEventArgs e)
        {
            myTextBox.Text = markupConverter.ConvertHtmlToXaml(myTextBox.Text);
            MessageBox.Show("Content Conversion Complete!");
        }

        public void convertXAMLToHTML(object sender, RoutedEventArgs e)
        {
            myTextBox2.Text = markupConverter.ConvertXamlToHtml(myTextBox2.Text);
            MessageBox.Show("Content Conversion Complete!");
        }

        public void convertRtfToHtml(object sender, RoutedEventArgs e)
        {
            myTextBox3.Text = markupConverter.ConvertRtfToHtml(myTextBox3.Text);
            MessageBox.Show("Content Conversion Complete!");
        }

        public void convertHtmlToRtf(object sender, RoutedEventArgs e)
        {
            myTextBox4.Text = markupConverter.ConvertHtmlToRtf(myTextBox4.Text);
            MessageBox.Show("Content Conversion Complete!");
        }

        public void convertXlsRTFtoHTML(object sender, RoutedEventArgs e)
        {
            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Open((myTextBox5.Text), Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing,
                                                       Type.Missing, Type.Missing);
                var ws = excelApp.Worksheets;
                var worksheet = (Worksheet)ws.get_Item("Sheet1");
                //Range range = ;
                //Range range = worksheet.UsedRange;
                object[,] values = new object[(Int32.Parse(upperRow5.Text) - Int32.Parse(lowerRow5.Text) + 1), (Int32.Parse(upperColumn5.Text) - Int32.Parse(lowerColumn5.Text) + 1)];//(object[,])range.Value2;

                for (int row = Int32.Parse(lowerRow5.Text); row <= Int32.Parse(upperRow5.Text); row++)    //row <= values.GetUpperBound(0)
                {
                    for (int column = Int32.Parse(lowerColumn5.Text); column <= Int32.Parse(upperColumn5.Text); column++ )
                    {
                        string cellName = convertCell(row, column);
                        string cellVal = Convert.ToString(worksheet.Range[cellName].Value);
                        try { 
                            cellVal = markupConverter.ConvertRtfToHtml(cellVal);
                        }
                        catch (Exception ex2) { }
                        //range.Cells.set_Item(row, column, cellVal);
                        worksheet.Range[cellName].Value = cellVal;
                    }
                }
                
                excelApp.ActiveWorkbook.SaveCopyAs(@"D:\Planlicht FOBs\RTF Converting\converted.xls");
                
                //excelApp.ActiveWorkbook.Save();
                excelApp.ActiveWorkbook.Close(true);
                excelApp.Quit();
           }
            catch (Exception ex)
            {//
                MessageBox.Show("Error while converting: " + ex.ToString());
            }

            MessageBox.Show("Excel Conversion Complete!");

        }

        public void copyXAML(object sender, RoutedEventArgs e)
        {
            myTextBox.SelectAll();
            myTextBox.Copy();
        }
        public void copyHTML(object sender, RoutedEventArgs e)
        {
            myTextBox2.SelectAll();
            myTextBox2.Copy();
        }

        public void copyHTML2(object sender, RoutedEventArgs e)
        {
            myTextBox3.SelectAll();
            myTextBox3.Copy();
        }

        public void copyRTF(object sender, RoutedEventArgs e)
        {
            myTextBox4.SelectAll();
            myTextBox4.Copy();
        }

        private string convertCell(int row, int column)
        {
            int dividend = column;
            string columnName = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName + row.ToString();
        }
    }
}