namespace Exchange.WebServices.NETCore.Tests.Folders;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class FolderPermissionsTests
{
    private readonly ExchangeProvider _provider;

    public FolderPermissionsTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Test]
    public async Task GetFolderPermissions()
    {
        var service = _provider.CreateTestService();
    }
}
