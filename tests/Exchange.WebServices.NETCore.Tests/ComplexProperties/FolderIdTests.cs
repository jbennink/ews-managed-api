using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.ComplexProperties;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class FolderIdTests
{
    private readonly ExchangeProvider _provider;


    public FolderIdTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Test]
    public async Task EqualityTests()
    {
        var a = new FolderId(WellKnownFolderName.AdminAuditLogs, new Mailbox());
        var b = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox());

        await Assert.That(a.Equals(b)).IsFalse();
        await Assert.That(a == b).IsFalse();
    }

    [Test]
    public async Task BrokenEquality()
    {
        var a = new FolderId(WellKnownFolderName.AdminAuditLogs);
        var b = new FolderId(WellKnownFolderName.ArchiveInbox);

        await Assert.That(a.Equals(b)).IsFalse();
    }

    [Test]
    public async Task EqualityTest()
    {
        var service = _provider.CreateTestService();

        var folders = (await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        )).ToList();

        await Assert.That(folders.Count > 2).IsTrue();

        var a = folders[0].Id;
        var b = folders[1].Id;

        Assert.NotNull(a);
        Assert.NotNull(b);


        var c = new FolderId(a.UniqueId);

        // ReSharper disable EqualExpressionComparison
        await Assert.That(a == a).IsTrue();
        await Assert.That(a.Equals(a)).IsTrue();
        await Assert.That(a == null).IsFalse();
        await Assert.That(null == a).IsFalse();

        await Assert.That(a.Equals(b)).IsFalse();
        await Assert.That(a == b).IsFalse();
        // ReSharper restore EqualExpressionComparison

        await Assert.That(a == c).IsTrue();
    }

    [Test]
    public async Task MailboxFolderIdEqualityTest()
    {
        var a = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox("hello@world.com"));
        var b = new FolderId(WellKnownFolderName.AdminAuditLogs, new Mailbox("world@hello.com"));
        var c = new FolderId(WellKnownFolderName.ArchiveInbox, new Mailbox("hello@world.com"));

        await Assert.That(a.Equals(a)).IsTrue();
        await Assert.That(a == b).IsFalse();

        await Assert.That(a == c).IsTrue();
    }
}
