# SCDashboard
Imports data from multiple stories to form a dashboard

### How to install from Visual Studio

Create a new C# console application with the name SCDashboard.

In the project folder, replace program.cs and app.config with the ones from this repo.

#Add References

Project -> Add References

System.Configuration

#Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Newtonsoft.Json

Install-Package SharpCloud.ClientAPI

Enter your Sharpcloud username, password, Sharpcloud URL, team, and dashboard story ID in the app.config file.
