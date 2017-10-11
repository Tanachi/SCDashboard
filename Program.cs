using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using System.Linq;
namespace SCDashboard
{
    class Program
    {
        static void Main(string[] args)
        {
            // Gets info from App.config
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var port = ConfigurationManager.AppSettings["port"];
            var storyID = ConfigurationManager.AppSettings["story"];
            var team = ConfigurationManager.AppSettings["team"];
            var cost = ConfigurationManager.AppSettings["cost"];
            var sc = new SharpCloudApi(userid, passwd);
            // Gets Story ID from URL
            string storyUrl = "";
            MatchCollection matchUrl = Regex.Matches(storyID, @"story\/(.+)\/view");
            string[] matchGroup = null;
            if (matchUrl.Count > 0)
            {
                matchGroup = matchUrl[0].ToString().Split('/');
                storyUrl = matchGroup[1];
            }
            else
            {
                storyUrl = storyID;
            }
            // Sets the team stories and the dash board story
            var teamBook = sc.StoriesTeam(team);
            var dashStory = sc.LoadStory(storyUrl);
            // Adds new attributes if story does not have it
            if (dashStory.Attribute_FindByName("Appropriated Budget") == null)
                dashStory.Attribute_Add("Appropriated Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            if (dashStory.Attribute_FindByName("RAG Status") == null)
                dashStory.Attribute_Add("RAG Status", SC.API.ComInterop.Models.Attribute.AttributeType.List);
            if (dashStory.Attribute_FindByName("New Requested Budget") == null)
                dashStory.Attribute_Add("New Requested Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);

            //Selected Roadmaps to insert into dashboard
            foreach(var teamStory in teamBook)
            {
                Story story = sc.LoadStory(teamStory.Id);
                if(story.Attribute_FindByName("Appropriated Budget") != null && story.Attribute_FindByName("New Requested Budget") != null)
                {
                    Console.WriteLine("Reading " + story.Name);
                    highCost(story, dashStory, false);
                }
                else
                {
                    Console.WriteLine("Budget not found in " + teamStory.Name); 
                }
            }
            // Goes through each roadmap to insert data into dashboard
            Console.WriteLine("Saving Dashboard");
            dashStory.Save();
        }
        static void highCost(Story story, Story newStory, bool isListed)
        {
            String catName = "";
            String newCat = "";
            // Finds the category that contains the project
            foreach (var cat in story.Categories)
            {
                var catLine = cat.Name.Split(' ');
                // Checks to see if last word is project, if so, adds the other words to be a category for the dashboard
                if (catLine[catLine.Length - 1] == "Projects")
                {
                    catName = cat.Name;
                    String[] noProject = new String[catLine.Length - 1];
                    Array.Copy(catLine, noProject, catLine.Length - 1);

                    newCat = String.Join(" ", noProject);
                    if (newStory.Category_FindByName(newCat) == null)
                        newStory.Category_AddNew(newCat);
                }
            }
            // Copies attribute data from team story to dashboard story
            foreach (var item in story.Items)
            {
                //if(item.getAttributeValueAsText("Attribute Check") !=null)
                //  isListed = true;
                double checkBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                if (item.Category.Name == catName && checkBudget > 1000000 || isListed == true)
                {
                    Item scItem = newStory.Item_AddNew(item.Name);
                    scItem.Category = newStory.Category_FindByName(newCat);
                    scItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                    scItem.DurationInDays = item.DurationInDays;
                    double appBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                    double newBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget"));
                    var RAG = item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status"));
                    scItem.SetAttributeValue(newStory.Attribute_FindByName("Appropriated Budget"), appBudget);
                    scItem.SetAttributeValue(newStory.Attribute_FindByName("New Requested Budget"), newBudget);
                    scItem.SetAttributeValue(newStory.Attribute_FindByName("RAG Status"), RAG);
                }
            }
            newStory.Save();
        }
    }
}
