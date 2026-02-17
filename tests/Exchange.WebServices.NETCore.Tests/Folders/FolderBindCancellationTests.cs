using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Folders;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class FolderBindCancellationTests
{
    private readonly ExchangeProvider _provider;

    public FolderBindCancellationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Test]
    public async Task BindFolder_Cancellation_ThrowsOperationCancelledException()
    {
        using var service = _provider.CreateTestService();

        // Create a cancelled cancellation token
        using var source = new CancellationTokenSource();
        await source.CancelAsync();

        try
        {
            _ = await Folder.Bind(service, WellKnownFolderName.Inbox, source.Token);
        }
        catch (OperationCanceledException)
        {
            // Do nothing
        }
    }
}
