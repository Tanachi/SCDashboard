using System;
using System.Configuration;
using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System.Text.RegularExpressions;
using System.Linq;
namespace SCBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load Data from App.config
            var userid = ConfigurationManager.AppSettings["user"];
            var passwd = ConfigurationManager.AppSettings["pass"];
            var team = ConfigurationManager.AppSettings["team"];
            var att = ConfigurationManager.AppSettings["att"];
            var getStory = ConfigurationManager.AppSettings["story"];
            var otherStory = ConfigurationManager.AppSettings["otherStory"];
            StoryLite[] teamBook = null;
            SharpCloudApi sc = null;
            // Login and get story data from Sharpcloud
            sc = new SharpCloudApi(userid, passwd);
            Story dashStory = sc.LoadStory(getStory);
            teamBook = sc.StoriesTeam(team);
            // Goes through each story in the team
            foreach (var teamStory in teamBook)
            {
                // Default case 
                String catName = "DME and O&M";
                String newCat = "DOH";
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
                            // 1st case where first first word is all letters followed by IT
                            if (storyLine[1] == "IT")
                            {
                                newCat = storyLine[0];
                            }
                            // 2nd case where 3rd word is IT and will add the other 2 words as the category
                            else if (storyLine.Length > 2 && storyLine[2] == "IT")
                            {
                                string Upperletter = storyLine[1].ToLower();
                                newCat = storyLine[0] + " " + storyLine[1][0] + Upperletter.Substring(1);
                            }
                            // 3rd case where Department is the first word and Dashboard is not in the 2nd word
                            else if (storyLine[0] == "Department")
                            {
                                if (storyLine[3][0].ToString() == "&")
                                    newCat = storyLine[2][0].ToString() + "U" + storyLine[4][0].ToString();
                            }
                        }
                        // Checks to see if the category is missing in dashboard
                        if (dashStory.Category_FindByName(newCat) == null)
                        {
                            dashStory.Category_AddNew(newCat);
                        }
                    }
                    // Goes through each item in the story
                    foreach (var item in story.Items)
                    {
                        // Checks to see if there's a bad item in the story
                        MatchCollection matchUrl = Regex.Matches(item.Name, @"Item \d+|(DELETE)|delete");
                        // checks for category if there's no category with projects
                        if (item.Category.Name == catName && matchUrl.Count == 0)
                        {
                            string check = item.GetAttributeValueAsText(story.Attribute_FindByName("Publish?"));
                            if(check != null && check != "" && check == "No")
                            {
                                Console.WriteLine("Found a No");
                                continue;
                            }
                            // Array filled with selected columns
                            string[] columns = att.Split(',');
                            // checks to see if item exists and will update the item if it does
                            Item newItem = null;
                            if (dashStory.Item_FindByName(item.Name) == null)
                                newItem = dashStory.Item_AddNew(item.Name);
                            else
                                newItem = dashStory.Item_FindByName(item.Name);
                            newItem.Category = dashStory.Category_FindByName(newCat);
                            // check to see if there is description
                            if (item.Description != null && item.Description != "")
                            {
                                newItem.Description = item.Description;
                            }
                            else
                            {
                                newItem.Description = "No Description Available";
                            }
                            // Add defeault attributes from items
                            newItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                            newItem.DurationInDays = item.DurationInDays;
                            // Add in attributes 
                            foreach (string attr in columns)
                            {
                                SC.API.ComInterop.Models.Attribute current = story.Attribute_FindByName(attr);
                                SC.API.ComInterop.Models.Attribute dashCurrent = dashStory.Attribute_FindByName(attr);
                                // Check to see if attribute is a number
                                if (current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.Numeric)
                                {
                                    if (item.GetAttributeValueAsDouble(current) != null)
                                    {
                                        newItem.SetAttributeValue(dashCurrent, item.GetAttributeValueAsDouble(current));
                                    }
                                }
                                // Checks to see if attribute is text
                                else if (current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.Text 
                                    || current.Type == SC.API.ComInterop.Models.Attribute.AttributeType.List)
                                {
                                    if (item.GetAttributeValueAsText(current) != null && item.GetAttributeValueAsText(current).Count() != 0)
                                    {
                                       
                                        newItem.SetAttributeValue(dashCurrent, item.GetAttributeValueAsText(current));
                                    }
                                }                       
                            }
                            // Add Tags to the story
                            foreach (var tag in item.Tags)
                            {
                                // Check to see if tag is in the story
                                ItemTag oldTag = null;
                                if (story.ItemTag_FindByName(tag.Text) != null)
                                    oldTag = story.ItemTag_FindByName(tag.Text);
                                ItemTag dashTag = null;
                                if (dashStory.ItemTag_FindByName(tag.Text) == null)
                                    dashTag = dashStory.ItemTag_AddNew(tag.Text);
                                else
                                    dashTag = dashStory.ItemTag_FindByName(tag.Text);
                                // check to see if tag has a group
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
    }
}