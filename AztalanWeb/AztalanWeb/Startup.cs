using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AztalanWeb.Startup))]
namespace AztalanWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
