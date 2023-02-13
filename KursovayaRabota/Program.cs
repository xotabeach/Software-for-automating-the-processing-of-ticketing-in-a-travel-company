using System;
using System.Globalization;

namespace KursovayaRabota
{
    class Program
    {
        static void Main(string[] args)
        {
            int found = 0;
            int id = 0;
            TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
            bool unlimitedCycle = true;
            int menuValue;
            string password = "2007";
            string surname;
            string menuExit;
            Console.WriteLine("##########################################################################################################################\n" +
                "# Туристическая фирма ООО 'TarasovTour'                                                                                   #\n" +
                "# г.Казань ул. Красносельская д.51                                                                                        #\n" +
                "#                                                                                                                         #\n" +
                "#                                                                                                                         #\n" +
                "##########################################################################################################################");
         while(unlimitedCycle)
            {

            
            Console.WriteLine("\n");
            Console.WriteLine(" Меню:");
            Console.WriteLine("(Для выбора нужного пункта, введите цифру, расположенную рядом с пунктом)");
            Console.WriteLine(" 1-Заполнить заявку");
            Console.WriteLine(" 2-Посмотреть статус заявки");
            Console.WriteLine(" 3-Изменить данные заявки");
            Console.WriteLine(" 4-Личный кабинет");
            Console.WriteLine(" 5-Посмотреть существующие маршруты");
            Console.WriteLine("---------------------------");
            Console.WriteLine(" 6-Раздел администрации");
            Console.WriteLine("---------------------------");
            Console.WriteLine(" 7-Выход");
            Console.Write(" Введеите номер пункта : ");
            string menuCheck = Console.ReadLine();
            if(!Int32.TryParse(menuCheck, out menuValue))
            {
                    Console.Clear();
                    Console.WriteLine("--------------------------------------- \n" +
                        "Ошибка! Введён неверный формат \n Попробуйте снова.\n" +
                        "---------------------------------------");
                
            }
            else
            {
                switch (menuValue)
                {
                    case 1:
                            Console.Clear();
                            Order neworder = new Order();
                            
                            Console.Write("Введите вашу фамилию: ");

                            surname = neworder.TakeClientSurname();
                            id  = neworder.CheckClient(surname);
                            neworder.AddNewOrder(surname, id);
                            Console.Clear();
                            
                            break;
                    case 2:
                        Order checkorder = new Order();
                            //string surname2 = checkorder.TakeClientSurname();
                            //checkorder.CheckClient(surname2);
                            checkorder.CheckOrder();
                            Console.Clear();
                            break;
                    case 3:

                        Order changekorder = new Order();
                            
                            int changedata = changekorder.TakeNumberofOrder();
                            bool check = changekorder.CheckToChange(changedata);
                            int index = changekorder.TakeNumberofChange(check);
                            changekorder.ChangeOrderData(index, changedata);
                            Console.Clear();
                            break;
                    case 4:
                            Order changekorder2 = new Order();
                            Clients lk = new Clients();
                            Console.Write("Введите вашу фамилию: ");
                            
                            surname = lk.TakeClientSurname();
                            id = lk.CheckClient(surname);
                            lk.LK(id);
                            
                            Console.Clear();
                            break;
                    case 5:
                            Console.Clear();
                            Console.WriteLine(" Список всех маршрутов, доступных на данный момент:");
                            IRoute allmarsh = new Route();
                            
                            allmarsh.PrintInfo();
                            Console.ReadKey();
                            Console.Clear();
                            break;
                    case 6:
                            bool passcycle = true;
                            
                            Console.Clear();
                            Console.Write("Введите пароль :");
                            string passchek = Console.ReadLine();
                            found = passchek.IndexOf(":");
                            passchek = passchek.Substring(found + 1);
                            while(passcycle)
                            {
                                if (passchek == password)
                                {
                                    passcycle = false;

                                    Console.WriteLine(" Список всех клиентов компании:");
                                    Clients test = new Clients();
                                    test.WriteList();
                                    test.PrintInfo();
                                    Console.ReadKey();
                                    Console.Clear();
                                }
                                else
                                {
                                    Console.WriteLine("Неверный пароль!");
                                }
                            }
                           
                            
                            break;
                    case 7:
                        Console.Write("Вы уверены что хотите завешить работу?( Y / N) : ");
                        menuExit = Console.ReadLine();
                            
                            found = menuExit.IndexOf(":");
                            menuExit = menuExit.Substring(found + 1);
                            menuExit = menuExit.ToUpper();
                        if(menuExit == "Y")
                        {
                              Console.Clear();
                              unlimitedCycle = false;
                                
                        }
                        else 
                        if(menuExit == "N")
                        {
                                Console.Clear();
                                break;
                        }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("--------------------------------------- \n" +
                        "Ошибка! Введён неверный формат \n Попробуйте снова.\n" +
                        "---------------------------------------");
                                goto case 7;
                            }
                        break;
                }
            }

            
            }
        }
        static void AllRight()
        {
            Console.WriteLine("-------------------------\n" +
                "-                       -\n" +
                "-       Thats Good!      -\n" +
                "-                       -\n" +
                "-------------------------");
        }
    }
}
