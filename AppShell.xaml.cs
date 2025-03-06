namespace ScoreCard;

public partial class AppShell : Shell
{
	public AppShell()
	{
		InitializeComponent();
        Routing.RegisterRoute("DetailedSales", typeof(Views.DetailedSalesPage));
    }
}