using System.Diagnostics;

using Exchange.WebServices.NETCore.Tests.Credentials;
using Exchange.WebServices.NETCore.Tests.Utility;

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Web.TokenCacheProviders;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Items;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class ItemOperationTests : ExchangeProvider
{
    private readonly ExchangeProvider _provider;

    public ItemOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Test]
    public async Task ItemSearchFilterTest()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var items = await service.FindItems(WellKnownFolderName.Inbox, filter, view);
        await Assert.That(items).IsNotEmpty();
    }

    [Test]
    public async Task ItemSuccessionTest()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var items = await service.FindItems(WellKnownFolderName.Inbox, filter, view);
        await Assert.That(items).IsNotEmpty();

        foreach (var item in items)
        {
            var mailItem = await Item.Bind(service, item.Id, [ItemSchema.MimeContent,]);

            Assert.NotNull(mailItem.MimeContent);
        }
    }

    [Test]
    public async Task FindItems_Cancelled_ThrowsOperationCancelledException()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var source = new CancellationTokenSource();
        await source.CancelAsync();

        try
        {
            await service.FindItems(WellKnownFolderName.Inbox, filter, view, token: source.Token);
        }
        catch (OperationCanceledException)
        {
            // Do nothing
        }
    }

    [Test]
    public async Task FindItems_Contact_Works()
    {
        var options = _provider.OutlookConnectionOptions;

        var service = new ExchangeService
        {
            Credentials = new TokenProvider(options, _provider.GetRequiredService<IMsalTokenCacheProvider>()),
            UseDefaultCredentials = false,
            AcceptGzipEncoding = true,
            Url = new Uri(options.Url),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, options.ImpersonationUpn),
            TraceEnabled = true,
            TraceListener = new EwsTraceListener(),
        };

        var view = new ItemView(100);

        var result = await service.FindItems(WellKnownFolderName.Contacts, view);

        Debugger.Break();
    }
}
