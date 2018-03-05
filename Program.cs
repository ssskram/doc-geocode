using System;
using System.IO;
using System.Data;
using System.Web;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using System.Collections.Specialized;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace sharepoint_geocode_documents
{
    class Program
    {
        HttpClient client = new HttpClient();
        static async Task Main()
        {
            Program run = new Program();
            await run.CodeMeta();
        }
        public async Task CodeMeta()
        {
            await refreshtoken();
            var token = refreshtoken().Result;
            // get list items
            var sharepointUrl = "https://cityofpittsburgh.sharepoint.com/sites/PublicSafety/ACC/_api/web/GetFolderByServerRelativeUrl('ScannedAdvises')/Files";
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            client.DefaultRequestHeaders.Authorization = 
            new AuthenticationHeaderValue ( "Bearer", token);
            string listitems = await client.GetStringAsync(sharepointUrl);
            dynamic items = JObject.Parse(listitems)["value"];

            char[] whitespace = {' ',' '};
            char[] period = {'.',' '};
            char[] brackets={'{','}',' '};
            char[] adv_char = {'A', 'D','V','a','d','v',' ' };
            char[] har_char = {'H', 'A','R','H','A','R',' ' };
            char[] pdf_char = {'P','D','F','p','d','f',' ' };
            char[] lat = {'"','l','a','t',':',' ' };
            
            int counter = 0;
            string filename = @"errors.csv";
            StreamWriter sw = new StreamWriter(filename);

            foreach (var item in items) 
            {
                var name = item.Name.ToString();

                // trim excess
                string adv_trimmed = name.TrimStart(adv_char);
                string har_trimmed = adv_trimmed.TrimStart(har_char);
                string pdf_trimmed = har_trimmed.TrimEnd(pdf_char);

                // encode name, generate string, and set to variable {link}
                var encodedName = System.Web.HttpUtility.UrlPathEncode(name);
                var link =
                    String.Format
                    ("https://cityofpittsburgh.sharepoint.com/sites/PublicSafety/ACC/ScannedAdvises/{0}",
                        encodedName); // 0

                // trim off date, clean, format, and set to variable {finaldate}
                string date = pdf_trimmed.Split(' ').First();
                string date_trimmed= date.TrimEnd(whitespace);
                string date_trimmed2 = date_trimmed.Replace(".", "-");
                string date_cleaned = "20" + date_trimmed2;
                DateTime finaldate; 
                bool parsed = DateTime.TryParseExact(date_cleaned, "yyyy-M-d", CultureInfo.InvariantCulture,
                            DateTimeStyles.AllowWhiteSpaces,
                            out finaldate);
                
                // trim off address, clean, format, and...
                string address = pdf_trimmed.Remove(0, pdf_trimmed.IndexOf(' ') + 1);
                string address_nowhitespace = address.TrimStart(whitespace);
                string address_trimmed = address_nowhitespace.TrimEnd(period);
                string address_formatted = 
                    String.Format 
                    ("{0}, Pittsburgh PA",
                        address_trimmed); // 0
                string address_encoded = address_formatted.Replace(" ", "+");

                if ((counter %2 == 0) || (counter == 0))
                {
                    var key1 = "AIzaSyD1CUnGKoU9ghcsXLJKEqY3CV-qxSn914E";
                    var geo_call =
                        String.Format 
                        ("https://maps.googleapis.com/maps/api/geocode/json?address={0}&key={1}",
                        address_encoded, // 0
                        key1); // 1
                    client.DefaultRequestHeaders.Clear();
                    string address_geocoded = await client.GetStringAsync(geo_call);

                    // check for geocoding success
                    dynamic status_check = JObject.Parse(address_geocoded)["status"];
                    if ( status_check == "OK")
                    {
                        // take response, and set formatted address to variable {formatted_address}
                        dynamic deseralize_4address = JsonConvert.DeserializeObject<dynamic>(address_geocoded)["results"][0];
                        string formatted_address = deseralize_4address.formatted_address.ToString();

                        // take response, format lat long to string, and set to variable {finalcoord}
                        dynamic deseralize_4coords = JsonConvert.DeserializeObject<dynamic>(address_geocoded)["results"][0]["geometry"];
                        string formatted_coords = deseralize_4coords.location.ToString();
                        var formatted_coords_nobrackets = formatted_coords.TrimEnd(brackets);
                        var formatted_coords_clean = formatted_coords_nobrackets.TrimStart(brackets);
                        string formatted_coords_lat = formatted_coords_clean.Remove(0, formatted_coords_clean.IndexOf(' ') + 1);
                        string formatted_coords_lat2 = formatted_coords_lat.TrimStart(lat);
                        string longitude_dirty = formatted_coords_lat2.Split(' ').Last();
                        string longitude = longitude_dirty.TrimEnd(whitespace);
                        string latitude = formatted_coords_lat2.Split(' ').FirstOrDefault();
                        var finalcoord =
                            String.Format 
                            ("({0} {1})",
                            latitude, // 0
                            longitude); // 1

                        await refreshtoken();
                        var updatedtoken = refreshtoken().Result;
                        // post data to new sharepoint list
                        var PUTsharepointUrl = "https://cityofpittsburgh.sharepoint.com/sites/PublicSafety/ACC/_api/web/lists/GetByTitle('GeocodedAdvises')/items";
                        client.DefaultRequestHeaders.Clear();
                        client.DefaultRequestHeaders.Authorization = 
                        new AuthenticationHeaderValue ("Bearer", updatedtoken);
                        client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                        client.DefaultRequestHeaders.Add("X-RequestDigest", "form digest value");
                        client.DefaultRequestHeaders.Add("X-HTTP-Method", "POST");
                        var dateformat = "MM/dd/yyyy HH:mm";
                        var otherjson = 
                            String.Format
                            ("{{'__metadata': {{ 'type': 'SP.Data.GeocodedAdvisesListItem' }}, 'Geo' : '{0}', 'Date' : '{1}', 'link' : '{2}', 'address' : '{3}' }}",
                                finalcoord, // 0
                                finaldate.ToString(dateformat), // 1
                                link, // 2
                                formatted_address); // 3
                                
                        client.DefaultRequestHeaders.Add("ContentLength", otherjson.Length.ToString());
                        StringContent stuff = new StringContent(otherjson);               
                        stuff.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                        HttpResponseMessage otherstuff = client.PostAsync(PUTsharepointUrl, stuff).Result;
                        otherstuff.EnsureSuccessStatusCode();
                        await otherstuff.Content.ReadAsStringAsync();
                        Console.WriteLine("SUCCESS " + formatted_address);   
                    }
                    else
                    {
                        sw.WriteLine(item.Name);
                    }  
                }
                else
                {
                    var key2 = "AIzaSyB20bjgHU5HBmaeyb1-f6s5dXlqMOxCnBo";
                    var geo_call =
                        String.Format 
                        ("https://maps.googleapis.com/maps/api/geocode/json?address={0}&key={1}",
                        address_encoded, // 0
                        key2); // 1
                    client.DefaultRequestHeaders.Clear();
                    string address_geocoded = await client.GetStringAsync(geo_call);

                    // check for geocoding success
                    dynamic status_check = JObject.Parse(address_geocoded)["status"];
                    if ( status_check == "OK")
                    {
                        // take response, and set formatted address to variable {formatted_address}
                        dynamic deseralize_4address = JsonConvert.DeserializeObject<dynamic>(address_geocoded)["results"][0];
                        string formatted_address = deseralize_4address.formatted_address.ToString();

                        // take response, format lat long to string, and set to variable {finalcoord}
                        dynamic deseralize_4coords = JsonConvert.DeserializeObject<dynamic>(address_geocoded)["results"][0]["geometry"];
                        string formatted_coords = deseralize_4coords.location.ToString();
                        var formatted_coords_nobrackets = formatted_coords.TrimEnd(brackets);
                        var formatted_coords_clean = formatted_coords_nobrackets.TrimStart(brackets);
                        string formatted_coords_lat = formatted_coords_clean.Remove(0, formatted_coords_clean.IndexOf(' ') + 1);
                        string formatted_coords_lat2 = formatted_coords_lat.TrimStart(lat);
                        string longitude_dirty = formatted_coords_lat2.Split(' ').Last();
                        string longitude = longitude_dirty.TrimEnd(whitespace);
                        string latitude = formatted_coords_lat2.Split(' ').FirstOrDefault();
                        var finalcoord =
                            String.Format 
                            ("({0} {1})",
                            latitude, // 0
                            longitude); // 1

                        await refreshtoken();
                        var updatedtoken = refreshtoken().Result;
                        // post data to new sharepoint list
                        var PUTsharepointUrl = "https://cityofpittsburgh.sharepoint.com/sites/PublicSafety/ACC/_api/web/lists/GetByTitle('GeocodedAdvises')/items";
                        client.DefaultRequestHeaders.Clear();
                        client.DefaultRequestHeaders.Authorization = 
                        new AuthenticationHeaderValue ("Bearer", updatedtoken);
                        client.DefaultRequestHeaders.Add("Accept", "application/json;odata=verbose");
                        client.DefaultRequestHeaders.Add("X-RequestDigest", "form digest value");
                        client.DefaultRequestHeaders.Add("X-HTTP-Method", "POST");
                        var dateformat = "MM/dd/yyyy HH:mm";
                        var otherjson = 
                            String.Format
                            ("{{'__metadata': {{ 'type': 'SP.Data.GeocodedAdvisesListItem' }}, 'Geo' : '{0}', 'Date' : '{1}', 'link' : '{2}', 'address' : '{3}' }}",
                                finalcoord, // 0
                                finaldate.ToString(dateformat), // 1
                                link, // 2
                                formatted_address); // 3
                                
                        client.DefaultRequestHeaders.Add("ContentLength", otherjson.Length.ToString());
                        StringContent stuff = new StringContent(otherjson);               
                        stuff.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                        HttpResponseMessage otherstuff = client.PostAsync(PUTsharepointUrl, stuff).Result;
                        otherstuff.EnsureSuccessStatusCode();
                        await otherstuff.Content.ReadAsStringAsync();
                        Console.WriteLine("SUCCESS " + formatted_address);   
                    }
                    else
                    {
                        sw.WriteLine(item.Name);
                    }  
                }

                counter++;
                
                }
            }
            public async Task<String> refreshtoken()
            {
                var MSurl = "https://accounts.accesscontrol.windows.net/f5f47917-c904-4368-9120-d327cf175591/tokens/OAuth/2";
                var clientid = "c0740d7b-dba6-425e-914f-859063d838fc%40f5f47917-c904-4368-9120-d327cf175591";
                var clientsecret = "kVcPCXVr50mTgObLxhgkV5d9b6CnYzFYlsxc9tdMN0w=";
                var refreshtoken = "IAAAAMsl76XQ7zdMfN0R192Cm7w5cUd7vOtIribo7QiEdX87oVoPrl-OnkLHgkXZuEHbkavAhUE2pcZaQBOCyyaaRGjyc7ZaXM0lFwfwSWNM3HFCFp7EBdmgHGW2HXIDS9xPuoUAZ-MKHuu63i430NitMzMdJvblsNhVkyu9IwwMjappKhPkebZBdNsypMnT8oncZiNfCSBit3YXug_bTH8qFC8Zjsd9pZKfeywrIcPfjAa1hIsZqgDzKJEiFL3HTqvbYwvglDjtAdUVK4sk7HO56LNmmmooiCHqMghNT8g_hWipNK06Y9sZewSRgy_KizidrSILXYfSdrI0gSBy_g_d340";
                var redirecturi = "https%3A%2F%2Flocalhost%2F";
                var SPresource = "00000003-0000-0ff1-ce00-000000000000%2Fcityofpittsburgh.sharepoint.com%40f5f47917-c904-4368-9120-d327cf175591";
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Add("Accept", "application/x-www-form-urlencoded");
                client.DefaultRequestHeaders.Add("X-HTTP-Method", "POST");

                var json =
                    String.Format 
                ("grant_type=refresh_token&client_id={0}&client_secret={1}&refresh_token={2}&redirect_uri={3}&resource={4}",
                    clientid, // 0
                    clientsecret, // 1
                    refreshtoken, // 2
                    redirecturi, // 3
                    SPresource); // 4

                client.DefaultRequestHeaders.Add("ContentLength", json.Length.ToString());

                StringContent strContent = new StringContent(json);               
                strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/x-www-form-urlencoded");
                HttpResponseMessage response = client.PostAsync(MSurl, strContent).Result;
                        
                response.EnsureSuccessStatusCode();
                var content = await response.Content.ReadAsStringAsync();
                dynamic results = JsonConvert.DeserializeObject<dynamic>(content);
                string token = results.access_token.ToString();
                return token;
            }
        }
    }
