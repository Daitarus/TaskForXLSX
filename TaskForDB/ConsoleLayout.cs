using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskForDB.Entities;
using TaskForDB.Repositories;

namespace TaskForDB
{
    internal class ConsoleLayout
    {
        private XLWorkbook workbook;
        private RepositoryProduct repositoryProduct;
        private RepositoryClient repositoryClient;
        private RepositoryOrder repositoryOrder;

        public ConsoleLayout()
        {
            Console.WriteLine("Добро пожаловать!");

            FileInfo fileXLSX = null;
            bool isInputCorrect = false;
            do
            {
                Console.Write("Введите путь к файлу с расширением .XLSX: ");
                string? path = Console.ReadLine();
                if (!String.IsNullOrEmpty(path))
                {
                    fileXLSX = new FileInfo(path);
                    isInputCorrect = fileXLSX.Exists && fileXLSX.Extension == ".XLSX";
                }
                if (!isInputCorrect)
                {
                    Console.WriteLine("Ошибка: Файла по указанному пути не существует или у него неправильное расширение!");
                }
            }
            while (!isInputCorrect);

            Console.WriteLine("Файл загружается...");
            workbook = new XLWorkbook(fileXLSX.FullName);
            repositoryProduct = new RepositoryProduct(workbook.Worksheet(1));
            repositoryClient = new RepositoryClient(workbook.Worksheet(2));
            repositoryOrder = new RepositoryOrder(workbook.Worksheet(3));
            Console.WriteLine("Файл загрузился!");
        }

        public void MainLayout()
        {
            List<string> listOfAction = new List<string>()
            {
                "Ввести имя товара узнать всю связанную информацию",
                "Изменит информацию о клиенте",
                "Узнать золотого клиента",
                "Завершить программу"
            };
            List<byte> listOfActionNumber = new List<byte>() {1, 2, 3, 4};
            
            bool isExit = false;
            do
            {
                bool isInputCorrect = false;
                byte numberOfAnswer = 0;
                do
                {
                    Console.Write("\n");
                    for (int i = 0; i < listOfAction.Count; i++)
                    {
                        Console.WriteLine("{0}. {1}", i + 1, listOfAction[i]);
                    }
                    Console.Write("Укажите номер действия: ");
                    string? stringAnswer = Console.ReadLine();
                    if (!String.IsNullOrEmpty(stringAnswer))
                    {
                        if (byte.TryParse(stringAnswer, out numberOfAnswer))
                            isInputCorrect = listOfActionNumber.Any(element => element == numberOfAnswer);
                    }
                }
                while (!isInputCorrect);

                switch (numberOfAnswer)
                {
                    case 1:
                        PrintProductInfoLayout();
                        break;
                    case 2:
                        ChageClient();
                        break;
                    case 3:
                        PrintGoldClient();
                        break;
                    case 4:
                        isExit = true;
                        break;
                    default:
                        break;

                }
            } while (!isExit);
        }

        public void PrintProductInfoLayout()
        {

            Console.Write("Введите имя продукта: ");
            string productName = Console.ReadLine();
            Product? product = repositoryProduct.GetForName(productName);
            if (product != null)
            {
                Console.WriteLine("Товар под номером: {0}", product.Id);
                List<Order> orders = repositoryOrder.SelectForProductId(product.Id);
                if (orders.Count > 0)
                {
                    foreach (Order order in orders)
                    {
                        Client? client = repositoryClient.GetForId(order.ClientId);
                        if(client!=null)
                        {
                            double sumPrice = order.Amount * product.Price;
                            Console.WriteLine("| - {0} (представитель {5}) заказал данный товар в количестве {1} {2}. Итоговая цена - {3}. Заказ был сделан - {4}", client.OrganizationName, order.Amount, product.NameUnit, sumPrice, order.DateOfCreation, client.ContactName);
                        }
                        else
                        {
                            Console.WriteLine("Ошибка: клиент по номером {0} отсутствует!", order.ClientId);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Данный товар ни кем не заказывался!");
                }
            }
            else
            {
                Console.WriteLine("Данного товара не существет!");
            }
        }

        public void PrintGoldClient()
        {
            List<Client> clients = repositoryClient.SelectAll();
            Client goldClient = clients[0];
            int maxOrders = 0;
            foreach (Client client in clients)
            {
                List<Order> ordersForClient = repositoryOrder.SelectForClientId(client.Id);
                int countOrderForClient = ordersForClient.Count();
                if(maxOrders<countOrderForClient)
                {
                    maxOrders = countOrderForClient;
                    goldClient = client;
                }
            }

            Console.WriteLine("На текущий момент золотым клиентом является - {0}", goldClient.OrganizationName);
        }

        public void ChageClient()
        {
            Console.Write("Введите имя организации: ");
            string organizationName = Console.ReadLine();
            Console.Write("Введите новое имя контактного лица: ");
            string newContactName = Console.ReadLine();

            Client client = repositoryClient.GetForOrganizationName(organizationName);
            if(client!=null)
            {
                repositoryClient.ChangeContactName(client.Id, newContactName);
                try
                {
                    workbook.Save();
                    Console.WriteLine("Изменение прошли учпешно!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка выполнения запроса!");
                }
            }
            else
            {
                Console.WriteLine("Клиент с таким именем не найден!");
            }
        }
    }
}
