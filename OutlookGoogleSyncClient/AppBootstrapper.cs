using System;
using System.Collections.Generic;
using System.Windows;
using Autofac;
using Caliburn.Micro;
using OutlookGoogleSyncClient.ViewModels;

namespace OutlookGoogleSyncClient
{
    public class AppBootstrapper : BootstrapperBase
    {
        private IContainer _container;

        public AppBootstrapper()
        {
            //Starts initialization sequence
            Initialize();
        }

        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            var settings = new Dictionary<string, object>
            {
                //{"Title", "Outlook Google Sync"},
            };

            //Start the main view
            DisplayRootViewFor<ShellViewModel>(settings);
        }

        protected override void Configure()
        {
            base.Configure();

            var builder = new ContainerBuilder();
            //builder.RegisterType<Services.Service1>()
            //    .AsImplementedInterfaces()
            //    .SingleInstance();

            //Register all ViewModels
            builder.RegisterAssemblyTypes(typeof(ShellViewModel).Assembly)
                .Where(t => t.Name.EndsWith("ViewModel"))
                .AsSelf();

            // register views, shells and dialogs, which are any type ends with 'View', 'Shell' or "Dialog'.
            builder.RegisterAssemblyTypes(typeof(ShellViewModel).Assembly)
                // must be a type with a name that ends with View
                .Where(type => type.Name.EndsWith("View") || type.Name.EndsWith("Shell") || type.Name.EndsWith("Dialog"))
                // registered as self
                .AsSelf()
                // always create a new one
                .InstancePerDependency();

            builder.RegisterType<WindowManager>().As<IWindowManager>().SingleInstance();

            _container = builder.Build();
        }

        /// <summary>
        /// After we configure the container, we need to tell Caliburn.Micro how to use it. 
        /// optionally used to supply property dependencies to instances of IResult that are executed by the Caliburn.Micro.
        /// </summary>
        /// <param name="instance"></param>
        protected override void BuildUp(object instance)
        {
            _container.InjectProperties(instance);
        }

        /// <summary>
        /// After we configure the container, we need to tell Caliburn.Micro how to use it. 
        /// That is the purpose of the three overrides that follow. “GetInstance” and “GetAllInstances” are required by the Caliburn.Micro.
        /// </summary>
        protected override object GetInstance(Type service, string key)
        {
            return _container.Resolve(service); //: _container.ResolveKeyed(key, service);
        }

        /// <summary>
        /// After we configure the container, we need to tell Caliburn.Micro how to use it. 
        /// That is the purpose of the three overrides that follow. “GetInstance” and “GetAllInstances” are required by the Caliburn.Micro.
        /// </summary>
        protected override IEnumerable<object> GetAllInstances(Type service)
        {
            var enumerableOfServiceType = typeof(IEnumerable<>).MakeGenericType(service);
            return (IEnumerable<object>)_container.Resolve(enumerableOfServiceType);
        }
    }
}
