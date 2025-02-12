using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Maui.Controls;


namespace ScoreCard.Resources
{
    public static class AppResources
    {
        public static ResourceDictionary CreateResources()
        {
            var resources = new ResourceDictionary();

            // Colors
            resources.Add("Primary", Color.FromHex("#FF4444"));
            resources.Add("Secondary", Color.FromHex("#2B2B2B"));
            resources.Add("Background", Colors.White);
            resources.Add("TextPrimary", Color.FromHex("#2B2B2B"));
            resources.Add("TextSecondary", Color.FromHex("#757575"));
            resources.Add("Success", Color.FromHex("#4CAF50"));
            resources.Add("Warning", Color.FromHex("#FFC107"));
            resources.Add("Error", Color.FromHex("#FF4444"));
            resources.Add("White", Colors.White);
            resources.Add("Black", Colors.Black);
            resources.Add("Gray100", Color.FromHex("#E1E1E1"));
            resources.Add("Gray200", Color.FromHex("#C8C8C8"));
            resources.Add("Gray300", Color.FromHex("#ACACAC"));
            resources.Add("Gray400", Color.FromHex("#919191"));
            resources.Add("Gray500", Color.FromHex("#6E6E6E"));
            resources.Add("Gray600", Color.FromHex("#404040"));
            resources.Add("Gray900", Color.FromHex("#212121"));
            resources.Add("Gray950", Color.FromHex("#141414"));

            // Styles
            resources.Add("DefaultLabel", new Style(typeof(Label))
            {
                Setters =
                {
                    new Setter { Property = Label.TextColorProperty, Value = resources["TextPrimary"] },
                    new Setter { Property = Label.FontFamilyProperty, Value = "OpenSansRegular" }
                }
            });

            resources.Add("CardFrame", new Style(typeof(Frame))
            {
                Setters =
                {
                    new Setter { Property = Frame.BackgroundColorProperty, Value = resources["Background"] },
                    new Setter { Property = Frame.BorderColorProperty, Value = resources["Gray100"] },
                    new Setter { Property = Frame.CornerRadiusProperty, Value = 8 },
                    new Setter { Property = Frame.PaddingProperty, Value = new Thickness(16) },
                    new Setter { Property = Frame.HasShadowProperty, Value = true }
                }
            });

            resources.Add("PrimaryButton", new Style(typeof(Button))
            {
                Setters =
                {
                    new Setter { Property = Button.BackgroundColorProperty, Value = resources["Primary"] },
                    new Setter { Property = Button.TextColorProperty, Value = Colors.White },
                    new Setter { Property = Button.FontFamilyProperty, Value = "OpenSansRegular" },
                    new Setter { Property = Button.CornerRadiusProperty, Value = 8 },
                    new Setter { Property = Button.PaddingProperty, Value = new Thickness(16, 8) }
                }
            });

            resources.Add("DashboardStatLabel", new Style(typeof(Label))
            {
                Setters =
                {
                    new Setter { Property = Label.FontSizeProperty, Value = 24 },
                    new Setter { Property = Label.FontFamilyProperty, Value = "OpenSansSemibold" },
                    new Setter { Property = Label.HorizontalOptionsProperty, Value = LayoutOptions.Start }
                }
            });

            resources.Add("DashboardStatTitle", new Style(typeof(Label))
            {
                Setters =
                {
                    new Setter { Property = Label.FontSizeProperty, Value = 14 },
                    new Setter { Property = Label.TextColorProperty, Value = resources["TextSecondary"] },
                    new Setter { Property = Label.FontFamilyProperty, Value = "OpenSansRegular" }
                }
            });

            return resources;
        }
    }
}