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
            var storyID = ConfigurationManager.AppSettings["story"];
            var team = ConfigurationManager.AppSettings["team"];
            var url = "https://my.sharpcloud.com";
            var enterprise = ConfigurationManager.AppSettings["enterprise"];
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
            Console.WriteLine("Logging in");
            var sc = new SharpCloudApi(userid, passwd, url);
            var teamBook = sc.StoriesTeam(team);
            var dashStory = sc.LoadStory(storyUrl);
            Console.WriteLine("Inserting Attributes");
            // Adds new attributes if story does not have it
            if (dashStory.Attribute_FindByName("Appropriated Budget") == null)
                dashStory.Attribute_Add("Appropriated Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            if (dashStory.Attribute_FindByName("RAG Status") == null)
                dashStory.Attribute_Add("RAG Status", SC.API.ComInterop.Models.Attribute.AttributeType.List);
            if (dashStory.Attribute_FindByName("New Requested Budget") == null)
                dashStory.Attribute_Add("New Requested Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            if (dashStory.Attribute_FindByName("Project Business Value") == null)
                dashStory.Attribute_Add("Project Business Value", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            if (dashStory.Attribute_FindByName("Project Dependencies/Assumptions/Risks") == null)
                dashStory.Attribute_Add("Project Dependencies/Assumptions/Risks", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
            if (dashStory.Attribute_FindByName("Total Spent to Date") == null)
                dashStory.Attribute_Add("Total Spent to Date", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            Story tagStory = sc.LoadStory(teamBook[0].Id);
            Console.WriteLine("Inserting Tags");
            // Add tags to new story
            foreach(var tag in tagStory.ItemTags)
            {
                if(dashStory.ItemTag_FindByName(tag.Name) == null)
                {
                    dashStory.ItemTag_AddNew(tag.Name, tag.Description, tag.Group);
                }
            }
            // Goes through entire team stories
            foreach (var teamStory in teamBook)
            {
                Story story = sc.LoadStory(teamStory.Id);
                // Does not filter based off budget if enterprise is false
                MatchCollection dashboard = Regex.Matches(story.Name, @"dashboard|Dashboard");
                if (enterprise == "false" && dashboard.Count == 0)
                {
                    Console.WriteLine("Regular Dashboard: " + story.Name);
                    highCost(story, dashStory, true, false);
                }
                // Filters based off budget
                /*
                else if(story.Attribute_FindByName("Appropriated Budget") != null && 
                    story.Attribute_FindByName("New Requested Budget") != null && dashboard.Count == 0)
                {
                    Console.WriteLine("Enterprise Dashboard: " + story.Name);
                    highCost(story, dashStory, false, false);
                }
                */
                // Story is not a roadmap with budget
                else
                {
                    Console.WriteLine("Budget not found " + teamStory.Name);
                }
            }
            
            // Goes through each roadmap to insert data into dashboard
            Console.WriteLine("Saving Story");
            dashStory.Save();
        }
        static void highCost(Story story, Story newStory, bool isBilled, bool cio)
        {
            String catName = "";
            String newCat = "empty";
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
                // Checks to see if there's a bad item in the story
                MatchCollection matchUrl = Regex.Matches(item.Name, @"Item \d+|(DELETE)");
                // checks for category if there's no category with projects
                if (newCat == "empty" && item.GetAllAttributeValues().Count > 8)
                {
                    newCat = item.Category.Name;
                    catName = item.Category.Name;
                    if (newStory.Category_FindByName(newCat) == null)
                    {
                        newStory.Category_AddNew(newCat);
                    }
                }
                // Final check to see if item has attribute values for category
                else if (newCat == "empty")
                {
                    if (item.GetAttributeValueAsText(story.Attribute_FindByName("Appropriated Budget")) != null
                        || item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value")) != null)
                    {
                        newCat = item.Category.Name;
                        catName = item.Category.Name;
                        if (newStory.Category_FindByName(newCat) == null)
                        {
                            newStory.Category_AddNew(newCat);
                        }
                    }
                }
                double checkBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                // Inserts item into dashboard
                if (item.Category.Name == catName)
                {
                    Item scItem = null;
                    if(newStory.Item_FindByName(item.Name) == null && item.Name != "" && matchUrl.Count == 0)
                    {
                        scItem = newStory.Item_AddNew(item.Name);
                        scItem.Category = newStory.Category_FindByName(newCat);
                        scItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                        scItem.DurationInDays = item.DurationInDays;
                        scItem.Description = item.Description;
                        // Get values from story
                        if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget")) != null)
                        {
                            double appBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("Appropriated Budget"), appBudget);
                        }
                        if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget")) != null)
                        {
                            double newBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("New Requested Budget"), newBudget);
                        }
                        if (item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status")) != null)
                        {
                            string RAG = item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("RAG Status"), RAG);
                        }
                        if (item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value")) != null)
                        {
                            string value = item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("Project Business Value"), value);
                        }
                        if (item.GetAttributeValueAsText(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks")) != null)
                        {
                            string risks = item.GetAttributeValueAsText(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("Project Dependencies/Assumptions/Risks"), risks);
                        }
                        if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks")) != null)
                        {
                            double total = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Total Spent to Date"));
                            scItem.SetAttributeValue(newStory.Attribute_FindByName("Total Spent to Date"), total);
                        }
                        foreach (var tag in item.Tags)
                        {
                            scItem.Tag_AddNew(tag.Text);
                        }
                    } 
                }
            }
        }
    }
}
