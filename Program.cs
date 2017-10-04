﻿using System;
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
            var URL = "https://my.sharpcloud.com";
            var sc = new SharpCloudApi(userid, passwd, URL);
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
            dashStory.Attribute_Add("Appropriated Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            dashStory.Attribute_Add("New Requested Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
            dashStory.Attribute_Add("RAG Status", SC.API.ComInterop.Models.Attribute.AttributeType.List);
            
            //Selected Roadmaps to insert into dashboard
            int[] maps = new int[4];
            maps[0] = 1;
            maps[1] = 3;
            maps[2] = 4;
            maps[3] = 5;
            // Goes through each roadmap to insert data into dashboard
            for(var i = 0; i < maps.Length; i++)
            {
                // Loads current story
                var storyGet = teamBook[maps[i]].Id;
                var story = sc.LoadStory(storyGet);
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
                        dashStory.Category_AddNew(newCat);
                    }
                }
                // Copies attribute data from team story to dashboard story
                foreach (var item in story.Items)
                {
                    if (item.Category.Name == catName)
                    {
                        Item scItem = dashStory.Item_AddNew(item.Name);
                        scItem.Category = dashStory.Category_FindByName(newCat);
                        scItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                        scItem.DurationInDays = item.DurationInDays;
                        double appBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                        double newBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget"));
                        var RAG = item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status"));
                        scItem.SetAttributeValue(dashStory.Attribute_FindByName("Appropriated Budget"), appBudget);
                        scItem.SetAttributeValue(dashStory.Attribute_FindByName("New Requested Budget"), newBudget);
                        scItem.SetAttributeValue(dashStory.Attribute_FindByName("RAG Status"), RAG);
                    }
                }
            }

            dashStory.Save();
        }
    }
}
