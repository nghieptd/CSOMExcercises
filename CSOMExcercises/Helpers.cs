using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace CSOMExcercises
{
    static class Helpers
    {
        public static async Task<TermCollection> GetAllTerms(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            if (taxonomySession == null)
            {
                throw new Exception("Cannot get taxonomy session");
            }

            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            ctx.Load(termStore, st => st.Groups.Include(g => g.Id, g => g.Name));
            await ctx.ExecuteQueryAsync();

            var termGroup = termStore.Groups.FirstOrDefault(group => group.Name == Constants.Taxonomy.TermGroupName);
            ctx.Load(termGroup);
            await ctx.ExecuteQueryAsync();
            if (termGroup == null)
            {
                throw new Exception("Missing group");
            }

            var termSetColl = termGroup.TermSets;
            var results = ctx.LoadQuery(termSetColl.Where(t => t.Name == Constants.Taxonomy.TermSetName));
            await ctx.ExecuteQueryAsync();

            var termSet = results.FirstOrDefault();
            if (termSet == null)
            {
                throw new Exception("Missing term set");
            }
            var terms = termSet.Terms;
            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();

            return terms;
        }
        public static async Task<Term> GetTermByEnum(ClientContext ctx, Constants.Taxonomy.TermsIndex term)
        {
            var terms = await GetAllTerms(ctx);
            return terms.FirstOrDefault(t => t.Name == Constants.Taxonomy.GetTermName(term));
        }
    }
}
