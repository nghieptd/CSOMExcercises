using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
                //await AddNewFieldAndMigrate(context);
                //await AddCitiesField(context);
                //await UpdateContentTypes(context);
                //await AddCitiesItems(context);
                //await AddDocumentLibrary(context);
                //await CreateDocumentHierachy(context);
                //await QueryDocument(context);
                await UploadFile(context);
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
                        Type='TaxonomyFieldType'
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
                var oListItem = list.AddItem(itemCreateInfo);
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
    Name='TestAuthor'
    Type='User'
    UserSelectionMode='PeopleOnly'
/>
", true, AddFieldOptions.AddFieldInternalNameHint);
            list.Update();
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
                // NOTE: The name "Author" already taken as internal field name
                item["TestAuthor"] = fieldUserValue;
                item.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
        public static async Task AddCitiesField(ClientContext ctx)
        {
            Console.WriteLine("Creating cities site column");

            Web rootWeb = ctx.Site.RootWeb;
            Field citiesField = rootWeb.Fields.AddFieldAsXml(@$"
                    <Field
                        DisplayName='{Constants.Columns.Cities.DisplayName}'
                        Name='{Constants.Columns.Cities.Name}'
                        Type='TaxonomyFieldTypeMulti'
                        Mult='TRUE'
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
            TaxonomyField tmField = ctx.CastTo<TaxonomyField>(citiesField);
            tmField.SspId = termStoreId;
            tmField.TermSetId = termSetId;
            tmField.TargetTemplate = String.Empty;
            tmField.AnchorId = Guid.Empty;
            tmField.Update();

            await ctx.ExecuteQueryAsync();
        }
        public static async Task UpdateContentTypes(ClientContext ctx)
        {
            Console.WriteLine("Updating content type with new field");

            var rootWeb = ctx.Site.RootWeb;
            ContentTypeCollection ctColl = rootWeb.ContentTypes;
            ctx.Load(ctColl);
            await ctx.ExecuteQueryAsync();

            var csomTestCt = ctColl.FirstOrDefault(ct => ct.Name == Constants.ContentType.Name);
            var citiesField = rootWeb.Fields.GetByInternalNameOrTitle(Constants.Columns.Cities.Name);
            csomTestCt.FieldLinks.Add(new FieldLinkCreationInformation { Field = citiesField });

            csomTestCt.Update(true);
            await ctx.ExecuteQueryAsync();
        }
        public static async Task AddCitiesItems(ClientContext ctx)
        {
            Console.WriteLine("Adding new items with cities");
            var data = new[]
            {
                (
                    "Test about with cities",
                    Constants.Taxonomy.TermsIndex.HoChiMinh,
                    new[]
                    {
                        Constants.Taxonomy.TermsIndex.HoChiMinh,
                        Constants.Taxonomy.TermsIndex.Stockholm
                    }
                ),
                (
                    "Test about with cities, part two",
                    Constants.Taxonomy.TermsIndex.HoChiMinh,
                    new[]
                    {
                        Constants.Taxonomy.TermsIndex.Stockholm
                    }
                ),
                (
                    "Test about with cities, final part",
                    Constants.Taxonomy.TermsIndex.Stockholm,
                    new[]
                    {
                        Constants.Taxonomy.TermsIndex.Stockholm,
                        Constants.Taxonomy.TermsIndex.HoChiMinh
                    }
                )
            };
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

            var list = ctx.Web.Lists.GetByTitle(Constants.List.Name);
            var tmpField = list.Fields.GetByInternalNameOrTitle(Constants.Columns.City.Name);
            var cityField = ctx.CastTo<TaxonomyField>(tmpField);
            tmpField = list.Fields.GetByInternalNameOrTitle(Constants.Columns.Cities.Name);
            var citiesField = ctx.CastTo<TaxonomyField>(tmpField);
            foreach (var item in data)
            {
                ListItemCreationInformation itemCreationInformation = new ListItemCreationInformation();
                var spItem = list.AddItem(itemCreationInformation);

                spItem["About"] = item.Item1;

                var cityTerm = terms.FirstOrDefault(t => t.Name == Constants.Taxonomy.GetTermName(item.Item2));
                if (cityTerm != null)
                {
                    cityField.SetFieldValueByTerm(spItem, cityTerm, 1033);
                }
                else
                {
                    Console.WriteLine("Missing city term. Skipping column");
                }

                var fieldColl = new TaxonomyFieldValueCollection(ctx, null, citiesField);
                foreach (var termName in item.Item3)
                {
                    var term = terms.FirstOrDefault(t => t.Name == Constants.Taxonomy.GetTermName(termName));
                    if (term != null)
                    {
                        fieldColl.PopulateFromLabelGuidPairs($"{term.Name}|{term.Id.ToString().ToLower()}");
                    }
                    else
                    {
                        Console.WriteLine("Missing city term. Skipping value");
                    }
                }
                citiesField.SetFieldValueByValueCollection(spItem, fieldColl);

                spItem.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
        public static async Task AddDocumentLibrary(ClientContext ctx)
        {
            ListCreationInformation creationInformation = new ListCreationInformation()
            {
                Title = Constants.Document.Name,
                TemplateType = (int) ListTemplateType.DocumentLibrary
            };

            var newLib = ctx.Web.Lists.Add(creationInformation);
            ctx.Load(newLib);
            var ctColl = ctx.Site.RootWeb.ContentTypes;
            ctx.Load(ctColl);
            await ctx.ExecuteQueryAsync();

            var ctCsomTest = ctColl.FirstOrDefault(ct => ct.Name == Constants.ContentType.Name);
            if (ctCsomTest == null)
            {
                throw new Exception("Missing content type");
            }

            newLib.ContentTypes.AddExistingContentType(ctCsomTest);
            newLib.Update();
            await ctx.ExecuteQueryAsync();
        }
        public static async Task CreateDocumentHierachy(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle(Constants.Document.Name);
            var rootFolder = list.RootFolder;

            // Add Folder 1
            var folder1 = rootFolder.Folders.Add("Folder 1");
            ctx.Load(folder1);
            await ctx.ExecuteQueryAsync();

            // Add Folder 2
            var folder2 = folder1.Folders.Add("Folder 2");
            ctx.Load(folder2);
            await ctx.ExecuteQueryAsync();
            // and 3 files
            for (int i = 0; i < 3; i++)
            {
                var sampleFileInfo = new FileCreationInformation()
                {
                    Url = $"file_1{i+1}.txt",
                    Content = Encoding.ASCII.GetBytes("Hello, World!")
                };
                var addedFile = folder2.Files.Add(sampleFileInfo);
                ctx.Load(addedFile);
                await ctx.ExecuteQueryAsync();

                var itemFile = addedFile.ListItemAllFields;
                itemFile["Title"] = $"Test file generation {i + 1}";
                itemFile["About"] = "Folder test";
                itemFile.Update();
            }
            await ctx.ExecuteQueryAsync();

            // ...and 2 more files, getting more elaborate
            var field = list.Fields.GetByInternalNameOrTitle(Constants.Columns.Cities.Name);
            var citiesField = ctx.CastTo<TaxonomyField>(field);
            var stockholmTerm = await Helpers.GetTermByEnum(ctx, Constants.Taxonomy.TermsIndex.Stockholm);
            var sampleFile2Info = new FileCreationInformation()
            {
                Url = "file_2.txt",
                Content = Encoding.ASCII.GetBytes("Hello there.")
            };
            var sampleFile3Info = new FileCreationInformation()
            {
                Url = "file_3.txt",
                Content = Encoding.ASCII.GetBytes("General Kenobi!")
            };

            var addedFile2 = folder2.Files.Add(sampleFile2Info);
            var addedFile3 = folder2.Files.Add(sampleFile3Info);
            ctx.Load(addedFile2);
            ctx.Load(addedFile3);
            await ctx.ExecuteQueryAsync();

            var itemFile2 = addedFile2.ListItemAllFields;
            var fieldColl = new TaxonomyFieldValueCollection(ctx, null, citiesField);
            fieldColl.PopulateFromLabelGuidPairs($"{stockholmTerm.Name}|{stockholmTerm.Id.ToString().ToLower()}");
            citiesField.SetFieldValueByValueCollection(itemFile2, fieldColl);
            itemFile2.Update();
            var itemFile3 = addedFile3.ListItemAllFields;
            fieldColl = new TaxonomyFieldValueCollection(ctx, null, citiesField);
            fieldColl.PopulateFromLabelGuidPairs($"{stockholmTerm.Name}|{stockholmTerm.Id.ToString().ToLower()}");
            citiesField.SetFieldValueByValueCollection(itemFile3, fieldColl);
            itemFile3.Update();
            await ctx.ExecuteQueryAsync();
        }
        public static async Task QueryDocument(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle(Constants.Document.Name);
            var rootFolder = list.RootFolder;
            ctx.Load(rootFolder);
            await ctx.ExecuteQueryAsync();

            var camlQuery = new CamlQuery();
            camlQuery.FolderServerRelativeUrl = $"{rootFolder.ServerRelativeUrl}/Folder 1/Folder 2";
            camlQuery.ViewXml = $@"
<View>
<Query>
    <Where>
        <Contains>
        <FieldRef Name='Cities' />
        <Value Type='Text'>Stockholm</Value>
    </Contains>
    </Where>
</Query>
</View>
";
            var collListItems = list.GetItems(camlQuery);
            ctx.Load(collListItems);
            await ctx.ExecuteQueryAsync();

            foreach (var item in collListItems)
            {
                Console.WriteLine($"ID: {item["ID"]} - Name: {item["FileLeafRef"]}");
            }
        }
        public static async Task UploadFile(ClientContext ctx)
        {
            // https://sharepoint.stackexchange.com/questions/121904/create-a-document-in-sharepoint-online-using-csom
            using (var stream = new MemoryStream())
            {
                using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
                {
                    document.AddMainDocumentPart();
                    document.MainDocumentPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Some content goes here")))));
                }

                var list = ctx.Web.Lists.GetByTitle(Constants.Document.Name);
                var rootFolder = list.RootFolder;

                // NOTE: Need to reset the stream before sending to SP
                stream.Seek(0, SeekOrigin.Begin);
                var fci = new FileCreationInformation()
                {
                    Url = "Document.docx",
                    Overwrite = true,
                    ContentStream = stream,
                };
                list.RootFolder.Files.Add(fci);

                await ctx.ExecuteQueryAsync();
            }
        }
    }
}
