using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CourseWork.StartMYup))]
namespace CourseWork
{
    public partial class StartMYup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
