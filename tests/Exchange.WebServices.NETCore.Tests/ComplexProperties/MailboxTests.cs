using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.ComplexProperties;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class MailboxTests
{
    private readonly ExchangeProvider _provider;


    public MailboxTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Test]
    public async Task EqualityTest()
    {
        var a = new Mailbox("hello@world.com");
        var b = new Mailbox("world@hello.com");
        var c = new Mailbox("hello@world.com");

        // ReSharper disable EqualExpressionComparison
        await Assert.That(a == a).IsTrue();
        await Assert.That(a != a).IsFalse();
        await Assert.That(a == null).IsFalse();
        await Assert.That(null == a).IsFalse();
        await Assert.That(a == b).IsFalse();
        await Assert.That(a != b).IsTrue();

        await Assert.That(a.Equals(b)).IsFalse();
        await Assert.That(a.Equals(null)).IsFalse();
        await Assert.That(Equals(a, a)).IsTrue();
        await Assert.That(Equals(a, b)).IsFalse();
        // ReSharper restore EqualExpressionComparison

        await Assert.That(a == c).IsTrue();
    }
}
