using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KursovayaRabota
{
    class Admin: Route
    {
        string[,] marshes; 

        public override void AddRoute()
        {
            
            

        }
        public string[] TakeData()
        {
            marshes =Marshbase2;
            string[] newmarsh = new string[marshes.GetLength(0)]; 
            int numberofmarsh = 0;
            string country = "";
            string city = "";
            string climate = "";
            string hotel = "";
            int longoftravel = 0;
            int cost = 0;
            return newmarsh;
        }
    }
}
