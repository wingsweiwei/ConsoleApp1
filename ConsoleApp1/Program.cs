namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //new SpirePdfTest().Run();
                //new OpenXmlTest().Run();
                //new TestWord2().Run();

                // new HostTest().Run();
                // new FreeSqlTest().Run();

                //new ExcelTest().Test();

                Console.WriteLine("End of program.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                Console.ReadKey();
            }
        }
    }
}
