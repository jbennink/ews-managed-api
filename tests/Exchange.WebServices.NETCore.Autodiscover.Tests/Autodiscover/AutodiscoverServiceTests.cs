using System.Diagnostics;

using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Autodiscover.Tests.Autodiscover;

[ClassDataSource<AutodiscoverProvider>(Shared = SharedType.PerClass)]
public class AutodiscoverServiceTests
{
    private readonly AutodiscoverProvider _provider;

    public AutodiscoverServiceTests(AutodiscoverProvider provider)
    {
        _provider = provider;
    }

    [Test]
    public async Task DiscoveryTest()
    {
        var options = _provider.ConnectionOptions.Value;

        var service = new AutodiscoverService
        {
            Credentials = new WebCredentials(options.UserName, options.Password),
        };


        var result = await service.GetUserSettings(
            options.UserName,
            UserSettingName.AlternateMailboxes,
            UserSettingName.ExternalEwsUrl
        );

        Debugger.Break();
    }

    [Test]
    public async Task DomainTest()
    {
        var options = _provider.ConnectionOptions.Value;


        var service = new AutodiscoverService
        {
        };

        var result = await service.GetDomainSettings(
            "outlook.com",
            ExchangeVersion.Exchange2007_SP1,
            DomainSettingName.ExternalEwsUrl
        );

        Debugger.Break();
    }

    [Test]
    public async Task LegacyDiscoveryTest()
    {
        var options = _provider.ConnectionOptions.Value;

        var service = new AutodiscoverService(ExchangeVersion.Exchange2007_SP1)
        {
            EnableScpLookup = false,
        };

        await service.AutodiscoverUrl(options.UserName, _ => true);
    }

    [Test]
    public async Task AutoDiscovery_SetTimeout_NoException()
    {
        var options = _provider.ConnectionOptions.Value;

        var service = new AutodiscoverService
        {
            Credentials = new WebCredentials(options.UserName, options.Password),
        };

        await service.GetUserSettings(options.UserName, UserSettingName.ExternalEwsUrl);
    }
}
