using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace KursovayaRabota
{
   
    
    public interface IRoute 
    {
        
        void PrintInfo();
        
    }
    
    

    class Route : IRoute 
    {
        protected int numberMarsh;
        protected int cost;
        protected string time_on;
        protected string climat;
        protected string hotel;
        protected string country;
        protected string city;
        protected string[,] marshBase;
        string text = "";
        protected List<string> baseMarsh = new List<string>();
        int k;
        int h;
        public string[,] Marshbase2
        {
            get { return marshBase; }
        }
        public int Marshbase{
            get { return marshBase.GetLength(0); }
            }
        public Route()
        {
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = true;
            var excelAppworkbookS = excelapp.Workbooks;
            var excelAppworkbook = excelapp.Workbooks.Open(@"C:\Users\taras\source\repos\KursovayaRabota\bin\Debug\net5.0\Kursovaya.xlsx ");
            var excelsheets = excelAppworkbook.Worksheets;
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
            Excel.Range xlRange = excelworksheet.UsedRange;
            marshBase = new string[xlRange.Count / 7, 7];
            for (int i = 2; i <= xlRange.Count / 7; i++)
            {
                for (char j = 'A'; j <= 'G'; j++)
                {

                    //Console.WriteLine(xlRange.Count);
                    var excelcells = excelworksheet.get_Range($"{j}{i}", Type.Missing);
                    string sStr = Convert.ToString(excelcells.Value2);

                    text += sStr + " ";
                marshBase[h, k] = sStr;
                    k++;

                }
                k = 0;
                baseMarsh.Add(text);
                h++;
                text = "";
            }
            excelapp.Quit();
        }
        public string InfoAboutMarshbyIndex(int index)
        {
            string text = "";
            for (int i = 0; i< marshBase.GetLength(1);i++)
                {
                switch (i)
                {
                    case 0:
                        numberMarsh = Convert.ToInt32(marshBase[index, 0]);
                        Console.WriteLine("Маршрут под номером "+ numberMarsh);
                        break;
                    case 1:
                        country = marshBase[index, 1];
                        
                        break;
                    case 2:
                        city = marshBase[index, 2];
                        break;
                    case 3:
                        climat = marshBase[index, 3];
                        break;
                    case 4:
                        hotel = marshBase[index, 4];
                        break;
                    case 5:
                        time_on = marshBase[index, 5];
                        break;
                    case 6:
                        cost = Convert.ToInt32(marshBase[index, 6]);
                        break;
                }
                
            }
            text = country + " " + city + " "+ climat + " "+ hotel + " "+ time_on + " " + cost ;
            return text;
        }
        
        void IRoute.PrintInfo()
        {

            string info = " ";
            for (int i = 0; i < marshBase.GetLength(0); i++)
            {
                for (int j = 0; j < marshBase.GetLength(1); j++)
                {
                    info += $" {marshBase[i, j]}  ";

                }
                Console.WriteLine($"\n Маршрут {i} {info}");
                info = "";
            }
            
        }
        
        public virtual void AddRoute()
        {

        }
        /*void IRoute.ChPref()
        {
            int number;
            Console.WriteLine("Выберите по какому условию, вы хотите выбрать путёвку. \n " +
                    "Или же вы можете воспользоваться нашей рекомендацией(Для этого нажмите цифру 5): ");
            Console.WriteLine("1- Ввести страну, которую бы вы хотели посетить \n" +
                "2- Ввести Продолжительность путевки, какая была бы для вас удобнее всего \n" +
                "3- Ввести примерную цену, за которую вы бы хотели увидеть путёвки" +
                "4- Ввести город, который вам интересен");
            string choose = Console.ReadLine();
            int chnum = choose.IndexOf(":");
            choose = choose.Substring(chnum + 1);
            if (Int32.TryParse(choose, out number))
            {
                switch (number)
                {
                    case 1:
                        break;
                    case 2:
                        break;
                    case 3:
                        break;
                    case 4:
                        break;
                    case 5:
                        break;
                }
            }
        }*/
        
        

    }
}
