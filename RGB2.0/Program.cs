using System;
using System.Drawing;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        string filePath = "C:/Documenti/RGB.xlsx";


        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Utenti");

            //Intestazione
            worksheet.Cell(1, 1).Value = "R";
            worksheet.Cell(1, 2).Value = "G";
            worksheet.Cell(1, 3).Value = "B";


            //Formattazione Intestazione
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true; //Grassetto
            headerRow.Style.Fill.BackgroundColor = XLColor.Black; //Colore Sfondo Cella
            headerRow.Style.Font.FontColor = XLColor.White; //Colore del Testo

            //Colori
            int nC;
            Console.WriteLine("Inserisci il numero di colori che vuoi avere");
            nC = Convert.ToInt16(Console.ReadLine());
            for (int i = 0; i < nC; i++)
            {
                //Rosso
                int rosso;
                do
                {
                    Console.WriteLine("COLORE " + (i + 1) + ": Inserisci il numero decimale per la colorazione del Rosso");
                    rosso = Convert.ToInt16(Console.ReadLine());
                } while (rosso < 0 || rosso > 255);
                worksheet.Cell(i + 3, 1).Value = rosso;

                //Verde
                int verde;
                do
                {
                    Console.WriteLine("COLORE " + (i + 1) + ": Inserisci il numero decimale per la colorazione del Verde");
                    verde = Convert.ToInt16(Console.ReadLine());
                } while (verde < 0 || verde > 255);
                worksheet.Cell(i + 3, 2).Value = verde;

                //Blu
                int blu;
                do
                {
                    Console.WriteLine("COLORE " + (i + 1) + ": Inserisci il numero decimale per la colorazione del Blu");
                    blu = Convert.ToInt16(Console.ReadLine());
                } while (blu < 0 || blu > 255);
                worksheet.Cell(i + 3, 3).Value = blu;

                //Colore Finale
                worksheet.Cell(i + 3, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(rosso, verde, blu);
            }


            //Testo
            worksheet.Cell(2, 6).Value = "Colori Ottenuti";


            //Colonna
            worksheet.Column(6).Width = 17;


            workbook.SaveAs(filePath);


        }

    }
}