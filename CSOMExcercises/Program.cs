using System;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

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

                /*await CreateList(context);
                await CreateTerms(context);*/
                await CreateSiteColumns(context);
            }
        }
        public static async Task CreateList(ClientContext ctx)
        {
            Console.WriteLine("Creating List");
            Web web = ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = "CSOM Test";
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;

            List list = web.Lists.Add(creationInfo);
            list.Description = "CSOM Test List";

            list.Update();

            await ctx.ExecuteQueryAsync();
        }
        public static async Task CreateTerms(ClientContext ctx)
        {
            Console.WriteLine("Creating Terms");
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            if (taxonomySession != null)
            {
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
                if (termStore != null)
                {
                    /*1033 is the LCID (Locale Identifier) for North America */
                    TermGroup newGroup = termStore.CreateGroup("CSOM Test", Guid.NewGuid());
                    TermSet cityNghiep = newGroup.CreateTermSet("city-Nghiep", Guid.NewGuid(), 1033);
                    cityNghiep.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());
                    cityNghiep.CreateTerm("Stockholm", 1033, Guid.NewGuid());

                    await ctx.ExecuteQueryAsync();
                }
            }
        }
        public static async Task CreateSiteColumns(ClientContext ctx)
        {
            Console.WriteLine("Creating 'About' and 'City' site columns");

            Web rootWeb = ctx.Site.RootWeb;
            rootWeb.Fields.AddFieldAsXml(@"
                    <Field 
                        DisplayName='About'
                        Name='About'
                        Type='Text'
                    />
                ", false, AddFieldOptions.AddFieldInternalNameHint);
            Field xmlTmField = rootWeb.Fields.AddFieldAsXml(@"
                    <Field
                        DisplayName='City'
                        Name='City'
                        Type='TaxonomyFieldTypeMulti'
                    />
                ", false, AddFieldOptions.AddFieldInternalNameHint);

            await ctx.ExecuteQueryAsync();

            /* Find information about term sets and update to taxonomy field */
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName("city-Nghiep", 1033);

            ctx.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            ctx.Load(termStore, ts => ts.Id);
            await ctx.ExecuteQueryAsync();

            var termStoreId = termStore.Id;
            var termSetId = termSets.FirstOrDefault().Id;

            /* Update taxonomy field */
            TaxonomyField tmField = ctx.CastTo<TaxonomyField>(xmlTmField);
            tmField.SspId = termStoreId;
            tmField.TermSetId = termSetId;
            tmField.TargetTemplate = String.Empty;
            tmField.AnchorId = Guid.Empty;
            tmField.Update();

            await ctx.ExecuteQueryAsync();
        }
    }
}
