namespace ScoreCard;

public partial class AppShell : Shell
{
	public AppShell()
	{
		InitializeComponent();
        Routing.RegisterRoute("DetailedAnalysis", typeof(Views.DetailedSalesPage));
    }
}