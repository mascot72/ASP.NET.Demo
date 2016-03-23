using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExcelWeb.Startup))]
namespace ExcelWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
