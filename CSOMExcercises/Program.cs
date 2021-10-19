using System;
using System.Collections.Generic;
using System.Linq;
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

                //await CreateList(context);
                //await CreateTerms(context);
                //await CreateSiteColumns(context);
                //await CreateAndAssignContentTypes(context);
                //await CreateListItems(context);
                //await UpdateDefaultValuesAndAddItems(context);
                //await QueryData(context);
                //await CreateListView(context);
                //await BatchUpdate(context);
                await AddNewFieldAndMigrate(context);
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
        /*public static async Task PrintContentTypes(ClientContext ctx)
        {
            ContentTypeCollection ctCollection = ctx.Web.ContentTypes;
            ctx.Load(ctCollection);
            await ctx.ExecuteQueryAsync();

            foreach (var contentType in ctCollection)
            {
                Console.WriteLine($"Content Type Name: {contentType.Name}");
            }
        }*/
        public static async Task CreateAndAssignContentTypes(ClientContext ctx)
        {
            Console.WriteLine("Create new Content Type");
            ContentTypeCollection ctCollection = ctx.Web.ContentTypes;
            ctx.Load(ctCollection);
            await ctx.ExecuteQueryAsync();

            var parentCt = ctCollection.FirstOrDefault(contentType => contentType.Name == "Item");
            /* Use default Item content type */
            Web rootWeb = ctx.Site.RootWeb;
            var ctCsomTest = rootWeb.ContentTypes.Add(new ContentTypeCreationInformation { Name = "CSOM Test Content Type", ParentContentType = parentCt });
            var titleField = ctCsomTest.Fields.GetByTitle("Title");
            ctx.Load(titleField);
            await ctx.ExecuteQueryAsync();
            /* Set Title parent column as hidden and not required */
            var titleRef = ctCsomTest.FieldLinks.GetById(titleField.Id);
            titleRef.Hidden = true;
            titleRef.Required = false;

            /* Add existing site columns */
            Field city = rootWeb.Fields.GetByInternalNameOrTitle("City");
            Field about = rootWeb.Fields.GetByInternalNameOrTitle("About");
            ctCsomTest.FieldLinks.Add(new FieldLinkCreationInformation { Field = city });
            ctCsomTest.FieldLinks.Add(new FieldLinkCreationInformation { Field = about });
            ctCsomTest.Update(true);
            await ctx.ExecuteQueryAsync();

            /* Assign to list as default type */
            Console.WriteLine("Assigning new content type to list");
            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            if (targetList != null)
            {
                targetList.ContentTypes.AddExistingContentType(ctCsomTest);
                targetList.Update();
                var contentTypesCol = targetList.ContentTypes;
                ctx.Load(contentTypesCol, col => col.Include(ct => ct.Name, ct => ct.Id));
                await ctx.ExecuteQueryAsync();

                var contentTypesOrder = new List<ContentTypeId>();
                foreach (var ct in contentTypesCol)
                {
                    Console.WriteLine(ct.Name);
                    if (ct.Name == "CSOM Test Content Type")
                    {
                        contentTypesOrder.Add(ct.Id);
                    }
                }
                targetList.RootFolder.UniqueContentTypeOrder = contentTypesOrder;
                targetList.RootFolder.Update();
                targetList.Update();
                await ctx.ExecuteQueryAsync();
            }
        }
        public static async Task CreateListItems(ClientContext ctx)
        {
            Console.WriteLine("Adding 5 new items to list");
            var data = new[]
            {
                ("Test about", "Stockholm"),
                ("Test about me and others", "Stockholm"),
                ("Another test about", "Ho Chi Minh"),
                ("Another test about which is longer than the rest", "Ho Chi Minh"),
                ("Another test about which is the longest about section", "Ho Chi Minh")
            };

            // Get terms
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            if (taxonomySession == null)
            {
                throw new Exception("Cannot get taxonomy session");
            }

            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            ctx.Load(termStore, st => st.Groups.Include(g => g.Id, g => g.Name));
            await ctx.ExecuteQueryAsync();

            var termGroup = termStore.Groups.FirstOrDefault(group => group.Name == "CSOM Test");
            ctx.Load(termGroup);
            await ctx.ExecuteQueryAsync();
            if (termGroup == null)
            {
                throw new Exception("Missing group");
            }

            var termSetColl = termGroup.TermSets;
            var results = ctx.LoadQuery(termSetColl.Where(t => t.Name == "city-Nghiep"));
            await ctx.ExecuteQueryAsync();

            var termSet = results.FirstOrDefault();
            if (termSet == null)
            {
                throw new Exception("Missing term set");
            }
            var terms = termSet.Terms;
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();

            // Add items from data to list
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            var field = list.Fields.GetByInternalNameOrTitle("City");
            var taxonomyField = ctx.CastTo<TaxonomyField>(field);
            foreach (var item in data)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = list.AddItem(itemCreateInfo);
                oListItem["About"] = item.Item1;

                // Find term by label
                var term = terms.FirstOrDefault(t => t.Name == item.Item2);
                if (term == null)
                {
                    Console.WriteLine("Term not found. Skipping");
                    continue;
                }
                taxonomyField.SetFieldValueByTerm(oListItem, term, 1033);

                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
        public static async Task UpdateDefaultValuesAndAddItems(ClientContext ctx)
        {
            // Load terms, list
            const string DefaultTermLabel = "Ho Chi Minh";
            string[] aboutDefaultData = new string[]
            {
                "Ho Chi Minh",
                "Stockholm"
            };
            string[] cityDefaultData = new string[]
            {
                "Testing default city",
                ""
            };
            Web rootWeb = ctx.Site.RootWeb;
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            if (taxonomySession == null)
            {
                throw new Exception("Cannot get taxonomy session");
            }

            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            ctx.Load(termStore, st => st.Groups.Include(g => g.Id, g => g.Name));
            await ctx.ExecuteQueryAsync();

            var termGroup = termStore.Groups.FirstOrDefault(group => group.Name == "CSOM Test");
            ctx.Load(termGroup);
            await ctx.ExecuteQueryAsync();
            if (termGroup == null)
            {
                throw new Exception("Missing group");
            }

            var termSetColl = termGroup.TermSets;
            var results = ctx.LoadQuery(termSetColl.Where(t => t.Name == "city-Nghiep"));
            await ctx.ExecuteQueryAsync();

            var termSet = results.FirstOrDefault();
            if (termSet == null)
            {
                throw new Exception("Missing term set");
            }
            var terms = termSet.Terms;
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();

            var txField = rootWeb.Fields.GetByInternalNameOrTitle("City");
            var aboutField = rootWeb.Fields.GetByInternalNameOrTitle("About");
            var taxonomyField = ctx.CastTo<TaxonomyField>(txField);
            var list = ctx.Web.Lists.GetByTitle("CSOM Test");

            // Update default value of "About" field and add 2 new items
            aboutField.DefaultValue = "about default";
            aboutField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();
            foreach (var item in aboutDefaultData)
            {
                var createItemInfo = new ListItemCreationInformation();
                var oListItem = list.AddItem(createItemInfo);

                var term = terms.FirstOrDefault(t => t.Name == item);
                if (term == null)
                {
                    Console.WriteLine("Term name not found. Skipping");
                    continue;
                }

                taxonomyField.SetFieldValueByTerm(oListItem, term, 1033);
                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();

            // Update default value of "City" and add 2 new items
            var defaultTerm = terms.FirstOrDefault(t => t.Name == DefaultTermLabel);
            if (defaultTerm == null)
            {
                throw new Exception("Default term not found");
            }
            // https://www.timmerman.it/index.php/setting-the-default-value-of-a-managed-metadata-column-from-csom/
            var defaultTxFieldValue = new TaxonomyFieldValue() { Label = defaultTerm.Name, TermGuid = defaultTerm.Id.ToString().ToLower(), WssId = -1 };
            var validatedValue = taxonomyField.GetValidatedString(defaultTxFieldValue);
            await ctx.ExecuteQueryAsync();

            taxonomyField.DefaultValue = validatedValue.Value;
            taxonomyField.UserCreated = false;
            taxonomyField.UpdateAndPushChanges(true);
            await ctx.ExecuteQueryAsync();
            foreach (var item in cityDefaultData)
            {
                var createItemInfo = new ListItemCreationInformation();
                var oListItem = list.AddItem(createItemInfo);
                oListItem["About"] = item;

                oListItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
        public static async Task QueryData(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("CSOM Test");
            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"
<View>
<Query>
    <Where>
        <Neq>
            <FieldRef Name='About' />
            <Value Type='Text'>about default</Value>
        </Neq>
    </Where>
</Query>
</View>
";
            ListItemCollection collListItem = list.GetItems(camlQuery);

            ctx.Load(collListItem);
            await ctx.ExecuteQueryAsync();

            foreach (var item in collListItem)
            {
                var city = item["City"] as TaxonomyFieldValue;
                Console.WriteLine($"ID: {item.Id}\nAbout: {item["About"]}\nCity: {city.Label}\n-----\n");
            }
        }
        public static async Task CreateListView(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("CSOM Test");

            ViewCollection viewCollection = list.Views;
            ctx.Load(viewCollection);
            await ctx.ExecuteQueryAsync();

            // To query against metadata
            // https://www.sharepointpals.com/post/retrieving-list-item-using-caml-query-against-taxonomy-field-in-sharepoint-2013/
            ViewCreationInformation viewCreationInformation = new ViewCreationInformation()
            {
                Title = "Test View",
                ViewTypeKind = ViewType.Html,
                Query = @"
<Where>
    <Contains>
        <FieldRef Name='City' />
        <Value Type='Text'>Ho Chi Minh</Value>
    </Contains>
</Where>
<OrderBy>
    <FieldRef Name='Created' Ascending='FALSE' />
</OrderBy>
",
                ViewFields = new string[] { "ID", "Title", "City", "About" }
            };
            var listView = viewCollection.Add(viewCreationInformation);
            await ctx.ExecuteQueryAsync();
            listView.Title = "Test View";
        }
        public static async Task BatchUpdate(ClientContext ctx, int batchNum = 2)
        {
            var list = ctx.Web.Lists.GetByTitle("CSOM Test");
            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"
<View>
<Query>
    <Where>
        <Eq>
            <FieldRef Name='About' />
            <Value Type='Text'>about default</Value>
        </Eq>
    </Where>
</Query>
<RowLimit>" + batchNum + @"</RowLimit>
</View>
";
            ListItemCollection collListItem = list.GetItems(camlQuery);

            ctx.Load(collListItem);
            await ctx.ExecuteQueryAsync();

            foreach (var item in collListItem)
            {
                item["About"] = "Update script";

                item.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
        public static async Task AddNewFieldAndMigrate(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("CSOM Test");
            var authorField = list.Fields.AddFieldAsXml(@"
<Field 
    DisplayName='Author'
    Name='Author'
    Type='User'
    UserSelectionMode='0'
/>
", true, AddFieldOptions.AddFieldInternalNameHint);
            await ctx.ExecuteQueryAsync();

            // Migrate all items with administrator user in author field
            var caml = CamlQuery.CreateAllItemsQuery();
            var itemColl = list.GetItems(caml);
            ctx.Load(itemColl);
            // Load user refs
            // https://social.msdn.microsoft.com/Forums/office/en-US/900b5143-f5b3-4fd5-a9ce-3e7d7c3ecfc1/csomjsom-operation-to-update-user-field-value?forum=sharepointdevelopment
            // https://stackoverflow.com/questions/31177333/ensureuser-using-email-address-in-sharepoint-client-object-model
            var result = Microsoft.SharePoint.Client.Utilities.Utility.ResolvePrincipal(
                ctx, ctx.Web, "nghiep@mianohr.onmicrosoft.com",
                Microsoft.SharePoint.Client.Utilities.PrincipalType.User,
                Microsoft.SharePoint.Client.Utilities.PrincipalSource.All,
                null, true
            );
            await ctx.ExecuteQueryAsync();
            if (result == null)
            {
                throw new Exception("User admin not found");
            }
            var user = ctx.Web.EnsureUser(result.Value.LoginName);
            ctx.Load(user);
            await ctx.ExecuteQueryAsync();
            var fieldUserValue = new FieldUserValue() { LookupId = user.Id };

            foreach (var item in itemColl)
            {
                item["Author"] = fieldUserValue;
                item.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
    }
}
