using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Graph.Core;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Xml;
using DocumentFormat.OpenXml;
using System.Linq;

namespace OpenXML
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            try
            {
                var siteUrl = @"https://pod3.sharepoint.com/sites/AuvenirDev_Local_-55c9bfd7-c7ee-4e2b-b7f1-f24a01e7b4ca/944e205f-9272-4e4d-aa6a-8d9702e5e0b4";

                using var content = await GetContentFileFromSPOAsync(siteUrl);
                var lstComments = GetListCommentsAsync(content);

                Console.WriteLine($"Find {lstComments.Count} comment.");

                if (lstComments.Count > 0)
                {
                    foreach (var cmt in lstComments)
                    {
                        Console.WriteLine(cmt);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);

            }

            Console.WriteLine("End.");

            Console.ReadKey();
        }

        public static List<string> GetListCommentsAsync(Stream stream)
        {
            //var fileName = @"D:\ConsoleTest\Syncfusion\Test.xlsx";
            using var spreadSheetDocument = SpreadsheetDocument.Open(stream, false);

            WorkbookPart workBookPart = spreadSheetDocument.WorkbookPart;
            List<string> lstComments = new List<string>();
            foreach (WorksheetPart sheet in workBookPart.WorksheetParts)
            {
                foreach (WorksheetCommentsPart commentsPart in sheet.GetPartsOfType<WorksheetCommentsPart>())
                {
                    foreach (Comment comment in commentsPart.Comments.CommentList)
                    {
                        if (comment != null)
                        {
                            lstComments.Add(comment.InnerText);
                        }
                    }
                }
            }

            return lstComments;
        }

        public static async Task<string> GetAccessTokenSPOAsync()
        {
            //var uri = new Uri(siteUrl);
            //var scope = $"{uri.Scheme}://{uri.Authority}/.default";
            var requestData = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>("client_id", "af33b34f-06de-4f7c-97fb-7a01e93064d4"),
                new KeyValuePair<string, string>("client_secret", ".WDibIBNEObf.bHl-o0N9xD6.DzN3zN5lK"),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                new KeyValuePair<string, string>("grant_type", "password"),
                new KeyValuePair<string, string>("username", "dev@pod3.onmicrosoft.com"),
                new KeyValuePair<string, string>("password", "12345678x@X")
            };

            var httpRequestMessage = new HttpRequestMessage
            {
                RequestUri = new Uri(@"https://login.microsoftonline.com/a36e627f-33eb-4ab9-9e01-653019cc7b55/oauth2/v2.0/token"),
                Method = HttpMethod.Post,
                Content = new FormUrlEncodedContent(requestData)
            };

            var httpClient = new HttpClient();
            var response = await httpClient.SendAsync(httpRequestMessage);
            var result = await response.Content.ReadAsStringAsync();

            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new Exception($"GetAccessTokenSPOAsync: Get access token failure: {result}");
            }

            var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
            return tokenResult.GetProperty("access_token").GetString();
        }

        public static async Task<Stream> GetContentFileFromSPOAsync(string siteUrl)
        {
            var token = await GetAccessTokenSPOAsync();

            var graphServiceClient = new GraphServiceClient(
                               new DelegateAuthenticationProvider(
                                   async (requestMessage) =>
                                   {
                                       requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                                   }));

            var uriSite = new Uri(siteUrl);
            var siteCollection = await graphServiceClient.Sites.GetByPath(uriSite.AbsolutePath, uriSite.Host).Request().GetAsync();

            var drive = graphServiceClient.Sites[siteCollection.Id].Drive.Root;
            var item = await drive.ItemWithPath("Engagement Request File/Test.xlsx").Content.Request().GetAsync();
            var comments = await drive.ItemWithPath("Engagement Request File/Test.xlsx").Workbook.Comments.Request().GetAsync();

            var workbookComment = new WorkbookComment
            {
                Content = "New comment from Graph api.",
                ContentType = "plain"
            };

            var comments1 = await drive.ItemWithPath("Engagement Request File/Test.xlsx").Workbook.Comments
                .Request()
                .AddAsync(workbookComment);

            foreach(var comment in comments)
            {
                var lstRelied = await drive
                    .ItemWithPath("Engagement Request File/Test.xlsx").Workbook.Comments[comment.Id].Replies
                    .Request()
                    .GetAsync();
            }

            return item;
        }

        //public static void InsertComments(WorksheetPart worksheetPart, Dictionary<string, string> commentsToAddDict)
        //{
        //    if (commentsToAddDict.Any())
        //    {
        //        string commentsVmlXml = string.Empty;

        //        // Create all the comment VML Shape XML
        //        foreach (var commentToAdd in commentsToAddDict)
        //        {
        //            commentsVmlXml += GetCommentVMLShapeXML(GetColumnName(commentToAdd.Key), GetRowIndex(commentToAdd.Key).ToString());
        //        }

        //        // The VMLDrawingPart should contain all the definitions for how to draw every comment shape for the worksheet
        //        VmlDrawingPart vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
        //        using (XmlTextWriter writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8))
        //        {

        //            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n" +
        //            "</o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n  path=\"m,l,21600r21600,l21600,xe\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n </v:shapetype>"
        //            + commentsVmlXml + "</xml>");
        //        }

        //        // Create the comment elements
        //        foreach (var commentToAdd in commentsToAddDict)
        //        {
        //            WorksheetCommentsPart worksheetCommentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();

        //            // We only want one legacy drawing element per worksheet for comments
        //            if (worksheetPart.Worksheet.Descendants<LegacyDrawing>().SingleOrDefault() == null)
        //            {
        //                string vmlPartId = worksheetPart.GetIdOfPart(vmlDrawingPart);
        //                LegacyDrawing legacyDrawing = new LegacyDrawing() { Id = vmlPartId };
        //                worksheetPart.Worksheet.Append(legacyDrawing);
        //            }

        //            Comments comments;
        //            bool appendComments = false;
        //            if (worksheetPart.WorksheetCommentsPart.Comments != null)
        //            {
        //                comments = worksheetPart.WorksheetCommentsPart.Comments;
        //            }
        //            else
        //            {
        //                comments = new Comments();
        //                appendComments = true;
        //            }

        //            // We only want one Author element per Comments element
        //            if (worksheetPart.WorksheetCommentsPart.Comments == null)
        //            {
        //                Authors authors = new Authors();
        //                Author author = new Author();
        //                author.Text = "Author Name";
        //                authors.Append(author);
        //                comments.Append(authors);
        //            }

        //            CommentList commentList;
        //            bool appendCommentList = false;
        //            if (worksheetPart.WorksheetCommentsPart.Comments != null &&
        //                worksheetPart.WorksheetCommentsPart.Comments.Descendants<CommentList>().SingleOrDefault() != null)
        //            {
        //                commentList = worksheetPart.WorksheetCommentsPart.Comments.Descendants<CommentList>().Single();
        //            }
        //            else
        //            {
        //                commentList = new CommentList();
        //                appendCommentList = true;
        //            }

        //            Comment comment = new Comment() { Reference = commentToAdd.Key, AuthorId = (UInt32Value)0U };

        //            CommentText commentTextElement = new CommentText();

        //            Run run = new Run();

        //            RunProperties runProperties = new RunProperties();
        //            Bold bold = new Bold();
        //            FontSize fontSize = new FontSize() { Val = 8D };
        //            Color color = new Color() { Indexed = (UInt32Value)81U };
        //            RunFont runFont = new RunFont() { Val = "Tahoma" };
        //            RunPropertyCharSet runPropertyCharSet = new RunPropertyCharSet() { Val = 1 };

        //            runProperties.Append(bold);
        //            runProperties.Append(fontSize);
        //            runProperties.Append(color);
        //            runProperties.Append(runFont);
        //            runProperties.Append(runPropertyCharSet);
        //            Text text = new Text();
        //            text.Text = commentToAdd.Value;

        //            run.Append(runProperties);
        //            run.Append(text);

        //            commentTextElement.Append(run);
        //            comment.Append(commentTextElement);
        //            commentList.Append(comment);

        //            // Only append the Comment List if this is the first time adding a comment
        //            if (appendCommentList)
        //            {
        //                comments.Append(commentList);
        //            }

        //            // Only append the Comments if this is the first time adding Comments
        //            if (appendComments)
        //            {
        //                worksheetCommentsPart.Comments = comments;
        //            }
        //        }
        //    }
        //}

        //private static string GetCommentVMLShapeXML(string columnName, string rowIndex)
        //{
        //    string commentVmlXml = string.Empty;

        //    // Parse the row index into an int so we can subtract one
        //    int commentRowIndex;
        //    if (int.TryParse(rowIndex, out commentRowIndex))
        //    {
        //        commentRowIndex -= 1;

        //        commentVmlXml = "<v:shape id=\"" + Guid.NewGuid().ToString().Replace("-", "") + "\" type=\"#_x0000_t202\" style=\'position:absolute;\r\n  margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;\r\n  visibility:hidden\' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\r\n  <v:fill color2=\"#ffffe1\"/>\r\n" +
        //        "<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\r\n  <v:path o:connecttype=\"none\"/>\r\n  <v:textbox style=\'mso-fit-shape-to-text:true'>\r\n   <div style=\'text-align:left\'></div>\r\n  </v:textbox>\r\n  <x:ClientData ObjectType=\"Note\">\r\n   <x:MoveWithCells/>\r\n" +
        //        "<x:SizeWithCells/>\r\n   <x:Anchor>\r\n" + GetAnchorCoordinatesForVMLCommentShape(columnName, rowIndex) + "</x:Anchor>\r\n   <x:AutoFill>False</x:AutoFill>\r\n   <x:Row>" + commentRowIndex + "</x:Row>\r\n   <x:Column>" + GetColumnIndexFromName(columnName) + "</x:Column>\r\n  </x:ClientData>\r\n </v:shape>";
        //    }

        //    return commentVmlXml;
        //}

        //private static string GetAnchorCoordinatesForVMLCommentShape(string columnName, string rowIndex)
        //{
        //    string coordinates = string.Empty;
        //    int startingRow = 0;
        //    int startingColumn = GetColumnIndexFromName(columnName).Value;

        //    // From (upper right coordinate of a rectangle)
        //    // [0] Left column
        //    // [1] Left column offset
        //    // [2] Left row
        //    // [3] Left row offset
        //    // To (bottom right coordinate of a rectangle)
        //    // [4] Right column
        //    // [5] Right column offset
        //    // [6] Right row
        //    // [7] Right row offset
        //    List<int> coordList = new List<int>(8) { 0, 0, 0, 0, 0, 0, 0, 0 };

        //    if (int.TryParse(rowIndex, out startingRow))
        //    {
        //        // Make the row be a zero based index
        //        startingRow -= 1;

        //        coordList[0] = startingColumn + 1; // If starting column is A, display shape in column B
        //        coordList[1] = 15;
        //        coordList[2] = startingRow;
        //        coordList[4] = startingColumn + 3; // If starting column is A, display shape till column D
        //        coordList[5] = 15;
        //        coordList[6] = startingRow + 3; // If starting row is 0, display 3 rows down to row 3

        //        // The row offsets change if the shape is defined in the first row
        //        if (startingRow == 0)
        //        {
        //            coordList[3] = 2;
        //            coordList[7] = 16;
        //        }
        //        else
        //        {
        //            coordList[3] = 10;
        //            coordList[7] = 4;
        //        }

        //        coordinates = string.Join(",", coordList.ConvertAll<string>(x => x.ToString()).ToArray());
        //    }

        //    return coordinates;
        //}

        ///// <summary>
        ///// Given just the column name (no row index), it will return the zero based column index.
        ///// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        ///// A length of three can be implemented when needed.
        ///// </summary>
        ///// <param name="columnName">Column Name (ie. A or AB)</param>
        ///// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        //public static int? GetColumnIndexFromName(string columnName)
        //{
        //    int? columnIndex = null;

        //    string[] colLetters = Regex.Split(columnName, "([A-Z]+)");
        //    colLetters = colLetters.Where(s => !string.IsNullOrEmpty(s)).ToArray();

        //    if (colLetters.Count() <= 2)
        //    {
        //        int index = 0;
        //        foreach (string col in colLetters)
        //        {
        //            List<char> col1 = colLetters.ElementAt(index).ToCharArray().ToList();
        //            int? indexValue = Letters.IndexOf(col1.ElementAt(index));

        //            if (indexValue != -1)
        //            {
        //                // The first letter of a two digit column needs some extra calculations
        //                if (index == 0 && colLetters.Count() == 2)
        //                {
        //                    columnIndex = columnIndex == null ? (indexValue + 1) * 26 : columnIndex + ((indexValue + 1) * 26);
        //                }
        //                else
        //                {
        //                    columnIndex = columnIndex == null ? indexValue : columnIndex + indexValue;
        //                }
        //            }

        //            index++;
        //        }
        //    }

        //    return columnIndex;
        //}
    }
}
