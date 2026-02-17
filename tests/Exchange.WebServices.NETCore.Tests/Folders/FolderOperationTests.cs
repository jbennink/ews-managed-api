using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Folders;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class FolderOperationTests
{
    private readonly ExchangeProvider _provider;

    public FolderOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Test]
    public async Task FindFoldersTest()
    {
        var service = _provider.CreateTestService();

        var folders = await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        );

        await Assert.That(folders).IsNotEmpty();
    }


    [Test]
    public async Task FolderBindTest()
    {
        var service = _provider.CreateTestService();

        var folder = await Folder.Bind(service, WellKnownFolderName.ArchiveRoot, PropertySet.FirstClassProperties);

        await Assert.That(folder).IsNotNull();
    }

    [Test]
    public async Task SyncFolderTest()
    {
        var service = _provider.CreateTestService();

        var calendarFolder = new FolderId(WellKnownFolderName.Calendar);
        var icc = await service.SyncFolderItems(
            calendarFolder,
            PropertySet.FirstClassProperties,
            null,
            123,
            SyncFolderItemsScope.NormalItems,
            "aaa"
        );
    }
}
