using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AsposeWordApplication.Startup))]
namespace AsposeWordApplication
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
