using System;
using System.Threading.Tasks;

namespace EWS_TestApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var app = new TestApp();
            await app.RunAsync();
        }
    }
}