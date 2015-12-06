using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Aztalan.Startup))]
namespace Aztalan
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
