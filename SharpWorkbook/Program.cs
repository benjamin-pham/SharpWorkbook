using SharpWorkbook;
using SharpWorkbook.ColumnCustoms;

List<User> users = User.GetUsers(1);

var workbookService = new WorkbookService();

var workbook = workbookService.CreateWorkbook();

workbook.AddSheet("Users", users);

workbook.SaveAs($"{Directory.GetCurrentDirectory()}\\Users.xlsx");

internal class User
{
    public string FirstName { get; set; } = null!;
    public string LastName { get; set; } = null!;
    public string Email { get; set; } = null!;
    public string PhoneNumber { get; set; } = null!;
    public DateTime DateOfBirth { get; set; }
    public string Gender { get; set; } = null!;
    public double Weight { get; set; }
    public double Height { get; set; }
    [ColumnIgnore]
    public string Country { get; set; } = null!;
    public string Province { get; set; } = null!;
    public string District { get; set; } = null!;
    public string Ward { get; set; } = null!;
    public string Street { get; set; } = null!;
    public static List<User> GetUsers(int count)
    {
        return Enumerable.Range(1, count).Select(_ => new User
        {
            FirstName = Guid.NewGuid().ToString(),
            LastName = Guid.NewGuid().ToString(),
            Email = Guid.NewGuid().ToString(),
            PhoneNumber = Guid.NewGuid().ToString(),
            DateOfBirth = DateTime.Now,
            Gender = Guid.NewGuid().ToString(),
            Height = 40,
            Weight = 70.5,
            Country = Guid.NewGuid().ToString(),
            Province = Guid.NewGuid().ToString(),
            District = Guid.NewGuid().ToString(),
            Ward = Guid.NewGuid().ToString(),
            Street = Guid.NewGuid().ToString(),
        }).ToList();
    }
}
