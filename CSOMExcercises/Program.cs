using System;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace CSOMExcercises
{
    class Program
    {
        public static IConfiguration GetConfiguration()
        {
            var builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", false, true);
            return builder.Build();
        }
        public static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
            {
                throw new ArgumentNullException("Missing password");
            }

            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            securePassword.MakeReadOnly();
            return securePassword;
        }
        public static async Task Main(string[] args)
        {
            var configurations = GetConfiguration();
            var sharepointConfigs = configurations.GetSection("SharePoint").Get<SharePointConfiguration>();
            /*Console.WriteLine(sharepointConfigs.SiteUri);*/
            
            // Note: The PnP Sites Core AuthenticationManager class also supports this
            using (var authenticationManager = new AuthenticationManager())
            using (var context = authenticationManager.GetContext(sharepointConfigs))
            {
                context.Load(context.Web, p => p.Title);
                await context.ExecuteQueryAsync();
                Console.WriteLine($"Title: {context.Web.Title}");
            }
        }
    }
}
