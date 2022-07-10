using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using GeoCoordinatePortable;

namespace TacoTake2
{
    public class Program
    {
        static void Main(string[] args)
        {

            //This is a complete recreate of the TacoParser project from the ground up for practice.

            //Phase 1: load excel document and populate locaiton list
            Console.WriteLine("Welcome! I will now parse an excel file for Taco Bell locations.");

            //Initializing excel document using Excel class
            Excel excel = new Excel();

            //Populating entries in a list of locations by parsing cell data from the excel document.
            string nextEntry = "";
            var currentRow = 0;
            var currentColumn = 0;
            List<TacoLocation> locations = new List<TacoLocation>();
            do
            {
                var location = new TacoLocation();
                location.latitude = Convert.ToDouble(excel.ReadCell(currentRow, currentColumn));
                //Console.WriteLine($"{excel.ReadCell(currentRow, currentColumn)}");
                currentColumn++;
                location.longitude = Convert.ToDouble(excel.ReadCell(currentRow, currentColumn));
                //Console.WriteLine($"{excel.ReadCell(currentRow, currentColumn)}");
                currentColumn++;
                location.name = Convert.ToString(excel.ReadCell(currentRow, currentColumn));
                //Console.WriteLine($"{excel.ReadCell(currentRow, currentColumn)}");
                currentRow++;
                currentColumn = 0;
                locations.Add(location);
                nextEntry = excel.ReadCell(currentRow, currentColumn);

            } while (nextEntry.Length > 1);

            //Phase 2: Identify furthest separated locations using GeoLocations
            Console.WriteLine($"Excel spreadsheet has been parsed, {locations.Count} locations have been identified.");

            var tacoBell1 = "";
            var tacoBell2 = "";
            double distance = 0;
            for (var i = 0; i < locations.Count; i++)
            {
                var locA = locations[i].name;

                var corA = new GeoCoordinate();
                corA.Latitude = locations[i].latitude;
                corA.Longitude = locations[i].longitude;

                for (var j = 0; j < locations.Count; j++)
                {
                    var locB = locations[j].name;

                    var corB = new GeoCoordinate();
                    corB.Latitude = locations[j].latitude;
                    corB.Longitude = locations[j].longitude;

                    if (corA.GetDistanceTo(corB) > distance)
                    {
                        distance = corA.GetDistanceTo(corB);
                        tacoBell1 = locA;
                        tacoBell2 = locB;
                    }
                }
            }
            //Convert to rounded miles for display
            distance = Math.Round(distance * 0.000621371);
            //Output
            Console.WriteLine($"After examination, I have determined that the {tacoBell1} and {tacoBell2} locations are furthest apart.");
            Console.WriteLine($"They are approximately {distance} miles apart.");
        }
    }
}
