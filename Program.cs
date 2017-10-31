using System;
using System.IO;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Net;
using System.Threading;
using OfficeOpenXml;
using System.Linq;
using System.Text;

namespace SCBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var fileLocation = Environment.CurrentDirectory.ToString() + "//Sheets";
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var team = ConfigurationManager.AppSettings["team"];
            var att = ConfigurationManager.AppSettings["att"];
            var getStory = ConfigurationManager.AppSettings["story"];
            var otherStory = ConfigurationManager.AppSettings["otherStory"];
            var csvLine = "Name,Description,Category,Start,Duration," + att + ",Tags.ETS Priorities, Tags.Governor Priorities";
            StoryLite[] teamBook = null;
            SharpCloudApi sc = null;
            // Login and get story data from Sharpcloud
            var filePath = System.IO.Directory.GetParent
             (System.IO.Directory.GetParent(Environment.CurrentDirectory)
             .ToString()).ToString() + "Hello.csv";
            sc = new SharpCloudApi(userid, passwd);
            Story dashStory = sc.LoadStory(getStory);
            Story checkStory = sc.LoadStory(otherStory);
            teamBook = sc.StoriesTeam(team);
	    // Goes through each story in the team
            foreach (var teamStory in teamBook)
            {
		//Default category is set to the exception
                String catName = "DME and O&M";
                String newCat = "DOH";
		// Check to see if story is a dashboard
                MatchCollection matchdash = Regex.Matches(teamStory.Name, @"Dashboard|dashboard");
                if (matchdash.Count == 0)
                {
                    Story story = sc.LoadStory(teamStory.Id);
                    foreach (var cat in story.Categories)
                    {
                        var catLine = cat.Name.Split(' ');
                        // Checks to see if last word is project, if so, adds the other words to be a category for the dashboard
                        if (catLine[catLine.Length - 1] == "Projects")
                        {
                            catName = cat.Name;
                            String[] noProject = new String[catLine.Length - 1];
                            var storyLine = story.Name.Split(' ');
                            if (storyLine[1] == "IT")
                            {
                                newCat = storyLine[0];
                            }
                            else if (storyLine.Length > 2 && storyLine[2] == "IT")
                            {
                                string Upperletter = storyLine[1].ToLower();
                                newCat = storyLine[0] + " " + storyLine[1][0] + Upperletter.Substring(1);
                            }
                            else if (storyLine[0] == "Department" && storyLine[1] != "Dashboard")
                            {
                                if (storyLine[3][0].ToString() == "&")
                                    newCat = storyLine[2][0].ToString() + "U" + storyLine[4][0].ToString();
                            }
                        }
                        if (dashStory.Category_FindByName(newCat) == null)
                        {
                            dashStory.Category_AddNew(newCat);
                        }
                    }
                    foreach (var item in story.Items)
                    {
                        var itemLine = "";
                        // Checks to see if there's a bad item in the story
                        MatchCollection matchUrl = Regex.Matches(item.Name, @"Item \d+|(DELETE)|delete");
                        // checks for category if there's no category with projects
                        // Inserts item into dashboard
                        if (item.Category.Name == catName && matchUrl.Count == 0)
                        {
                            string[] columns = att.Split(',');
                            Item newItem = null;
                            if (dashStory.Item_FindByName(item.Name) == null)
                                newItem = dashStory.Item_AddNew(item.Name);
                            else
                                newItem = dashStory.Item_FindByName(item.Name);
                            newItem.Category = dashStory.Category_FindByName(newCat);
                            if (item.Description != null && item.Description != "")
                            {
                                newItem.Description = item.Description;
                            }
                            else
                            {
                                newItem.Description = "No Description Available";
                            }
                            newItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                            newItem.DurationInDays = item.DurationInDays;
			    //Goes through the list of attributes in the config file and adds to items
                            foreach (string attr in columns)
                            {
                                SC.API.ComInterop.Models.Attribute current = story.Attribute_FindByName(attr);
                                SC.API.ComInterop.Models.Attribute dashCurrent = dashStory.Attribute_FindByName(attr);
                                if (current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.Numeric)
                                {
                                    if (item.GetAttributeValueAsDouble(current) != null)
                                    {
                                        newItem.SetAttributeValue(dashCurrent, item.GetAttributeValueAsDouble(current));
                                    }
                                    else
                                    {
                                        newItem.SetAttributeValue(dashCurrent, 0);
                                    }
                                }
                                else if (current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.Text || current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.List)
                                {
                                    if (item.GetAttributeValueAsText(current) != null && item.GetAttributeValueAsText(current).Count() != 0)
                                    {
                                        newItem.SetAttributeValue(dashCurrent, item.GetAttributeValueAsText(current));
                                    }
                                }
				// Adds tags to item
                                foreach (var tag in item.Tags)
                                {
                                    ItemTag oldTag = null;
                                    if (story.ItemTag_FindByName(tag.Text) != null)
                                        oldTag = story.ItemTag_FindByName(tag.Text);
                                    ItemTag dashTag = null;
                                    if (dashStory.ItemTag_FindByName(tag.Text) == null)
                                        dashTag = dashStory.ItemTag_AddNew(tag.Text);
                                    else
                                        dashTag = dashStory.ItemTag_FindByName(tag.Text);
                                    if (oldTag != null && oldTag.Group != "")
                                        dashTag.Group = oldTag.Group;
                                    if (newItem.Tag_FindByName(tag.Text) == null)
                                    {
                                        newItem.Tag_AddNew(dashTag);
                                    }
                                }
                            }

                        }
                    }
                }
            }
            dashStory.Save();
        }
    }
}