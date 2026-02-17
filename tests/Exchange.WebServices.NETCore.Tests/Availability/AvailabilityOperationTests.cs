using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Availability;

[ClassDataSource<ExchangeProvider>(Shared = SharedType.PerClass)]
public class AvailabilityOperationTests
{
    private readonly ExchangeProvider _provider;


    public AvailabilityOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }


    [Test]
    public async Task GetRoomListTest()
    {
        var service = _provider.CreateTestService();

        var rooms = await service.GetRoomLists();

        await Assert.That(rooms).IsEmpty();
    }
}
