using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Data.OleDb;
using Newtonsoft.Json;
using System.Data;
using System.Net.Http;
using System.Windows.Forms;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using System.Reflection;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace CompSci
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //global variables (all of these are required to be global)
        private ToastNotification toast;
        bool databaseloaded = false;
        public string connectionstring = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        private Dictionary<string, Dictionary<string, int>> graph;
        public DataTable Teachers = new DataTable();
        private TimeMonitor time;



        public MainWindow()
        {
            InitializeComponent();
            toast = new ToastNotification();
            initialisetimetable();
            customiseboxes();
        }


        private void initialisetimetable() //show user what the timetable will look like
        {
            DataTable timetable = new DataTable();
            timetable.Columns.Add("Time", typeof(string));
            timetable.Columns.Add("Monday", typeof(string));
            timetable.Columns.Add("Tuesday", typeof(string));
            timetable.Columns.Add("Wednesday", typeof(string));
            timetable.Columns.Add("Thursday", typeof(string));
            timetable.Columns.Add("Friday", typeof(string));
            timetable.Rows.Add(new TimeOnly(9, 00).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(10, 00).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(11, 00).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(11, 20).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(12, 20).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(12, 50).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(13, 20).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(13, 50).ToString("HH:mm"), "", "", "", "", "");
            timetable.Rows.Add(new TimeOnly(14, 40).ToString("HH:mm"), "", "", "", "", "");
            Timetable.ItemsSource = timetable.DefaultView;
            Timetable.IsReadOnly = true;
            info.Content = "Select database to use timetable feature";
        }

        private void customiseboxes()
        {
            foreach (PropertyInfo property in typeof(Colors).GetProperties(BindingFlags.Static | BindingFlags.Public)) //gets every colour available in c#
            {
                string colorName = property.Name;
                System.Windows.Media.Color colour = (System.Windows.Media.Color)property.GetValue(null);

                ComboBoxItem item = new ComboBoxItem //add the item to the dropdown box with background as colour and text as black/white depending on the contrast
                {
                    Content = colorName,
                    Background = new SolidColorBrush(colour),
                    Foreground = new SolidColorBrush(GetContrastingColor(colour))
                };
                backbox.Items.Add(item);

                ComboBoxItem item2 = new ComboBoxItem //cant add item to both boxes so have to redo the process
                {
                    Content = colorName,
                    Background = new SolidColorBrush(colour),
                    Foreground = new SolidColorBrush(GetContrastingColor(colour))
                };
                Notiftextcolour.Items.Add(item2);
                ComboBoxItem item3 = new ComboBoxItem
                {
                    Content = colorName,
                    Background = new SolidColorBrush(colour),
                    Foreground = new SolidColorBrush(GetContrastingColor(colour))
                };
                notifbackcolour.Items.Add(item3);

            }
        }

        private System.Windows.Media.Color GetContrastingColor(System.Windows.Media.Color colour) // takes the values for the current value and deduces if it will have better contrast against white or black
        {
            double brightness = (colour.R * 0.299 + colour.G * 0.587 + colour.B * 0.114) / 255;
            return brightness > 0.5 ? Colors.Black : Colors.White; // if brightness is bigger than 0.5, return black (else return white)
        }


        //SQL commands
        //----------
        //---------


        private void fillboxes()
        {
            start.Items.Add(""); //adds null values at the top in case the user wants to clear current selections
            end.Items.Add("");
            blocked.Items.Add("");
            teacherdropdown.Items.Add("");
            roomdropdown.Items.Add("");
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionstring)) //fills the datagrid with every teacher and their corresponding initials
                {
                    connection.Open();
                    string query = "SELECT Initials, RoomID FROM TeacherRooms";
                    OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(excecuteQuery);
                    adapter.Fill(Teachers);
                    TeacherDisplay.ItemsSource = Teachers.DefaultView;
                }
                foreach (DataRow row in Teachers.Rows) //fills out the dropdown boxes for teachers and rooms
                {
                    teacherdropdown.Items.Add(row[0]);
                    roomdropdown.Items.Add(row[1]);
                }
                using (OleDbConnection connection = new OleDbConnection(connectionstring)) //new query since there are rooms which dont appear in the previous query
                {
                    connection.Open();
                    string query = $"SELECT RoomID FROM Rooms";
                    OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                    using (OleDbDataReader reader = excecuteQuery.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            start.Items.Add(reader.GetString(0));
                            end.Items.Add(reader.GetString(0));
                            blocked.Items.Add(reader.GetString(0));
                            stop1.Items.Add(reader.GetString(0));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                toast.ShowNotification("Error: " + ex.Message);
            }
        }

        private DataTable loadtimetable()
        {
            DataTable loadedtt = new DataTable();
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionstring))
                {
                    connection.Open();
                    string query = "SELECT Time, Monday, Tuesday, Wednesday, Thursday, Friday FROM Timetable";
                    OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(excecuteQuery);
                    adapter.Fill(loadedtt);
                    Timetable.ItemsSource = loadedtt.DefaultView;
                }
                foreach (DataRow row in loadedtt.Rows)
                {
                    if (TimeOnly.TryParse(row["Time"].ToString(), out TimeOnly curenttime))//returns true or false depending on if the time within the database is of a valid TimeOnly format
                    {
                        row["Time"] = curenttime.ToString("HH:mm"); // only hours and minutes are needed for the datatable
                    }
                }
                Timetable.Columns[0].IsReadOnly = true; //dont allow the user to change the time column since it is constant for everyone
            }
            catch (Exception ex)
            {
                toast.ShowNotification("Error: " + ex.Message);
            }
            return loadedtt;

        }

        public (rooms, rooms, string) Corridorkey(string StartC, string EndC, string blocked)
        {

            StartC = StartC.ToUpper();
            EndC = EndC.ToUpper();
            (rooms Start, rooms End, string blockedC) Rooms = (new rooms(new string(""), new string(""), new string("")), new rooms(new string(""), new string(""), new string("")), new string(""));
            if (string.IsNullOrEmpty(blocked)) // prevent a potential exception for having a null value in one of the boxes
            {
                Rooms.blockedC = " ";
            }
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionstring))
                {
                    connection.Open();
                    string query = $"SELECT Rooms.RoomID, Rooms.Corridor, Corridors.Floor,Corridors.BuildingID FROM Rooms, Corridors WHERE (Rooms.RoomID = \"{StartC}\" OR Rooms.RoomID = \"{EndC}\" OR Rooms.RoomID = \"{blocked}\") AND (Rooms.Corridor = Corridors.CorridorName)"; // "\" key used to escape the speech marks (which are required for SQL)
                    OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                    using (OleDbDataReader reader = excecuteQuery.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader["RoomID"].ToString() == StartC) //fill in the objects for the start and end rooms
                            {
                                Rooms.Start.corridor = reader["Corridor"].ToString();
                                Rooms.Start.floor = reader["Floor"].ToString();
                                Rooms.Start.building = reader["BuildingID"].ToString();
                            }
                            else if (reader["RoomID"].ToString() == EndC)
                            {
                                Rooms.End.corridor = reader["Corridor"].ToString();
                                Rooms.End.floor = reader["Floor"].ToString();
                                Rooms.End.building = reader["BuildingID"].ToString();
                            }
                            else
                            {
                                Rooms.blockedC = reader["Corridor"].ToString(); //no need for full information here - only need to avoid the corridor
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                toast.ShowNotification("Error: " + ex.Message);
            }
            return Rooms;
        }

        private Dictionary<string, Dictionary<string, int>> CreateDbGraph()
        {
            Dictionary<string, Dictionary<string, int>> graph = new Dictionary<string, Dictionary<string, int>>();
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionstring))
                {
                    connection.Open();
                    string query = "SELECT Corridor1, Corridor2, Distance FROM CorridorConnections";
                    OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                    using (OleDbDataReader reader = excecuteQuery.ExecuteReader())
                    {
                        while (reader.Read()) //for every row in the returned data
                        {
                            string node1 = reader["Corridor1"].ToString();
                            string node2 = reader["Corridor2"].ToString();
                            int distance = int.Parse(reader["Distance"].ToString());
                            if (!graph.ContainsKey(node1)) //adds the room which isnt in the dictionary to it (and gives it its own dictionary to start adding connections to)
                            {
                                graph[node1] = new Dictionary<string, int>();
                            }
                            if (!graph.ContainsKey(node2))
                            {
                                graph[node2] = new Dictionary<string, int>();
                            }
                            graph[node1][node2] = distance; //ensures the graph is bidirectional (distance from A->B is equal to distance from B->A)
                            graph[node2][node1] = distance;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                toast.ShowNotification("Error: " + ex.Message);
            }
            return graph;
        }


        //A* algorithm
        //-------------
        //------------





        public List<string> Astar(string Start, string End, string blocked)
        {
            int totalGn = 0;
            List<(int, string)> openList = new List<(int Gn, string node)>();
            List<string> nopath = new List<string> { "No valid path for these inputs" };
            (rooms Start, rooms End, string blockedC) Rooms = Corridorkey(Start, End, blocked);
            if (string.IsNullOrEmpty(Rooms.Start.corridor) || string.IsNullOrEmpty(Rooms.End.corridor))//if either of the corridors for the input room is empty it means there wasnt one and the room isnt in the map
            {
                toast.ShowNotification("Rooms specified doesnt exist, try again");
                return nopath;
            }
            Start = Rooms.Start.corridor;
            End = Rooms.End.corridor;
            Dictionary<string, int> Gn = new Dictionary<string, int>();
            Dictionary<string, string> prevNode = new Dictionary<string, string>();
            foreach (string node in graph.Keys)
            {
                Gn[node] = int.MaxValue; //sets g(n) values at max (infinity) to show that they are unexplored
            }
            Gn[Start] = 0;
            openList = InitiateOpenList(0, Start, openList); //adds start node to open list with distance 0
            try
            {
                while (openList.Count > 0) //explores open list (due to graph setup once list is empty all possible nodes/paths will have been explored)
                {
                    (int currentFn, string currentNode) distlow = GnLow(openList);
                    if (distlow.currentNode == End)
                    {
                        return makepath(prevNode, distlow.currentNode); //if reached destination (prevents exploration of every possible path)
                    }
                    foreach (KeyValuePair<string, int> neighbour in graph[distlow.currentNode])//for each neighbour to the current node
                    {
                        totalGn = Gn[distlow.currentNode] + neighbour.Value;
                        if (totalGn < Gn[neighbour.Key]) //if the new cost is lower than the old cost
                        {
                            prevNode[neighbour.Key] = distlow.currentNode;
                            Gn[neighbour.Key] = totalGn;
                            int Fn = totalGn + Hn(Rooms.End, neighbour.Key, Rooms.blockedC); //calculate estimated cost
                            openList = InitiateOpenList(Fn, neighbour.Key, openList); //updates open list in preparation for next node exploration
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                toast.ShowNotification("Error: " + ex.Message);
            }
            return nopath;
        }

        public int Hn(rooms end, string current, string blocked)
        {
            if (current == blocked)
            {
                return 5000; //tells A* to never go down this route
            }
            else
            {
                rooms currentinfo = new rooms("", "", ""); //create a temporary object which will store the info for the current node
                try
                {
                    using (OleDbConnection connection = new OleDbConnection(connectionstring))
                    {
                        connection.Open();
                        string query = $"SELECT CorridorName, Floor, BuildingID FROM Corridors WHERE CorridorName = \"{current}\""; //get the info for only the currently looked at node
                        OleDbCommand excecuteQuery = new OleDbCommand(query, connection);
                        using (OleDbDataReader reader = excecuteQuery.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                currentinfo.corridor = reader["CorridorName"].ToString();
                                currentinfo.floor = reader["Floor"].ToString();
                                currentinfo.building = reader["BuildingID"].ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    toast.ShowNotification("Error: " + ex.Message);
                }
                if (end.building == currentinfo.building)
                {
                    if (end.floor == currentinfo.floor)
                    {
                        return 0;
                    }
                    else
                    {
                        return 1;
                    }
                }
                else
                {
                    return 10;
                }
            }
        }

        public static List<(int Fn, string node)> InitiateOpenList(int Gn, string node, List<(int Gn, string node)> openList)
        {
            openList.Add((Gn, node));
            openList.Sort((a, b) => a.Gn == b.Gn ? a.node.CompareTo(b.node) : a.Gn.CompareTo(b.Gn)); //sorts the list by order of g(n)
            return openList;
        }

        public static (int Fn, string node) GnLow(List<(int Gn, string node)> openList)
        {
            (int, string) lowest = openList[0];
            openList.RemoveAt(0); //removes node with lowest Gn from open list (shows it has been explored)
            return lowest;
        }

        public static List<string> makepath(Dictionary<string, string> prevNode, string currentNode)
        {
            List<string> path = new List<string> { currentNode };
            while (prevNode.ContainsKey(currentNode))
            {
                currentNode = prevNode[currentNode];
                path.Add(currentNode);
            }
            path.Reverse(); //currently the path is going from end -> start this reverses it
            return path; //completed path
        }


        //API requests
        //----------
        //-----------

        public async Task APIRequest((string latitude, string longitude) location) //Async so the other tasks can happen while this is running (prevents the whole program stopping if no response acheived)
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage //packages the request with URL and APIkeys ready to be sent off
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri($"https://meteostat.p.rapidapi.com/point/hourly?lat={location.latitude}&&lon={location.longitude}&start={today}&end={today}&alt=24&tz=Europe%2FLondon"),
                //Paramaterised API URI^^
                Headers =
                {
                    {
                        "x-rapidapi-key", "acb6abfc5emsh5fc6dfdbbf75c5bp1a226cjsn99be44419645"//unique API key
                    },
                    {
                        "x-rapidapi-host", "meteostat.p.rapidapi.com" //API host web address
                    },
                },
            };
            using (HttpResponseMessage response = await client.SendAsync(request)) //sends the request
            {
                response.EnsureSuccessStatusCode(); //typically code is 2xx
                string ResponseBody = await response.Content.ReadAsStringAsync();
                WeatherResponse weatherResponse = JsonConvert.DeserializeObject<WeatherResponse>(ResponseBody); //needs to be converted to objects before it can be used
                FormatData(weatherResponse.Data);
            }
        }

        public async Task LocationFinder(string postcode) //ensures other tasks can still run while this happens
        {

            (string latitude, string longitude) packedloc = ("", "");
            postcode = postcode.ToUpper(); //the api requires the parameter in all caps
            postcode = postcode.Replace(" ", "");
            int incodestart = postcode.Length - 3; //this is where the inward code starts
            string postcodeend = postcode.Substring(incodestart, 3); //inward code
            string postcodefront = postcode.Substring(0, incodestart); //outward code
            HttpClient client = new HttpClient();
            HttpRequestMessage request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri($"https://uk-postcode-to-latitude-and-longitude.p.rapidapi.com/?postcode={postcodefront}%20{postcodeend}"), //specific parameterized URI
                Headers =
                {
                    {
                        "x-rapidapi-key", "acb6abfc5emsh5fc6dfdbbf75c5bp1a226cjsn99be44419645"
                    },
                    {
                        "x-rapidapi-host", "uk-postcode-to-latitude-and-longitude.p.rapidapi.com"
                    },
                },
            };
            using (HttpResponseMessage response = await client.SendAsync(request))
            {
                response.EnsureSuccessStatusCode();
                string ResponseBody = await response.Content.ReadAsStringAsync();
                initresponse location = JsonConvert.DeserializeObject<initresponse>(ResponseBody);
                if (location.status == 1)
                {
                    packedloc = (location.output[0].Lat, location.output[0].Long);
                }
                else
                {
                    packedloc = ("51.633985", "0.422817");
                    toast.ShowNotification("No returned Data, used default location");
                }
                APIRequest(packedloc); //sends this data off the the weather API to retreive and display that data
            }
        }

        static StringBuilder compiledata(WeatherData data)
        {
            StringBuilder compileddata = new StringBuilder();
            compileddata.AppendLine($"Time: {data.Time}");
            compileddata.AppendLine($"Temperature: {data.Temperature}°C");
            compileddata.AppendLine($"Wind Speed: {data.WindSpeed} km/h");
            compileddata.AppendLine($"Precipitation: {data.Precipitation} mm");
            return compileddata;
        }

        public void FormatData(List<WeatherData> weatherData)
        {
            List<StringBuilder> compileddata = new List<StringBuilder>();
            foreach (WeatherData data in weatherData)
            {
                compileddata.Add(compiledata(data));
            }
            if (compileddata.Count > 15) //prevents exceptions if for some reason not enough data is given
            {
                displayweather(compileddata);
            }
            toast.ShowNotification("Data compiled successfully");
        }

        public void displayweather(List<StringBuilder> compileddata)
        {
            weather0.Items.Add(compileddata[9]);
            weather1.Items.Add(compileddata[10]);
            weather2.Items.Add(compileddata[11]);
            weather3.Items.Add(compileddata[12]);
            weather4.Items.Add(compileddata[13]);
            weather5.Items.Add(compileddata[14]);
            weather6.Items.Add(compileddata[15]);
            weather7.Items.Add(compileddata[16]);
        }

        public void clearweather()
        {
            weather0.Items.Clear();
            weather1.Items.Clear();
            weather2.Items.Clear();
            weather3.Items.Clear();
            weather4.Items.Clear();
            weather5.Items.Clear();
            weather6.Items.Clear();
            weather7.Items.Clear();
        }

        private bool giveoptions(string message)
        {
            MessageBoxResult result = System.Windows.MessageBox.Show(
                $"{message}", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question);
            return result == MessageBoxResult.Yes;
        }


        //classes
        //--------
        //--------


        public class initresponse
        {
            [JsonProperty("status")]
            public int status { get; set; }
            [JsonProperty("output")]
            public List<LatLong> output { get; set; }
        }

        public class LatLong
        {
            [JsonProperty("postcode")]
            public string postcode { get; set; }
            [JsonProperty("latitude")]
            public string Lat { get; set; }

            [JsonProperty("longitude")]
            public string Long { get; set; }
        }

        public class WeatherResponse
        {
            [JsonProperty("data")] //everything contained within the main body of the response (cuts off all unnecessary info)
            public List<WeatherData> Data { get; set; } //uses the weatherdata structure and maps the Json string onto it
        }

        public class WeatherData //extracts the wanted data and converts it into an easy to follow structure
        {
            [JsonProperty("time")] //Json propterties are strings which appear before the ":" in the raw response
            public string Time { get; set; } //isolates the specific value for that property and assigns it to a variable
            [JsonProperty("temp")]
            public double Temperature { get; set; }
            [JsonProperty("wspd")]
            public double WindSpeed { get; set; }
            [JsonProperty("prcp")]
            public double Precipitation { get; set; }
        }

        public class rooms
        {
            public string floor;
            public string building;
            public string corridor;
            public rooms(string floor, string building, string corridor)
            {
                this.floor = floor;
                this.building = building;
                this.corridor = corridor;
            }
        }

        public class TimeMonitor
        {
            private readonly DataTable scheduleTable;
            private readonly DispatcherTimer timer;
            private ToastNotification toast;
            public TimeMonitor(DataTable timeTable)
            {
                scheduleTable = timeTable;
                toast = new ToastNotification();
                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMinutes(1) //checks once per minute from when it is called (sets tick speed)
                };
                timer.Tick += CheckTimes;
            }

            private void CheckTimes(object sender, EventArgs e)
            {
                DateTime today = DateTime.Now;
                String currentTime = today.ToString("HH:mm");
                String currentDay = today.DayOfWeek.ToString();
                foreach (DataRow row in scheduleTable.Rows)
                {
                    TimeOnly actualtime = TimeOnly.Parse(row[0].ToString());
                    TimeOnly nextlesson = actualtime.AddMinutes(-5); //gets the time 5 minutes prior to the actual lesson time
                    string scheduleTime = nextlesson.ToString();
                    if ((int)today.DayOfWeek>=0 && (int)today.DayOfWeek<=5) //checks to see if its a weekday otherwise exceptions occur for referncing objects which dont appear in the datatable
                    {
                        if (currentTime == scheduleTime)
                        {
                            int dayColumnIndex = scheduleTable.Columns[currentDay].Ordinal;
                            string cellData = row[dayColumnIndex].ToString();
                            if (!string.IsNullOrEmpty(cellData))
                            {
                                toast.ShowNotification($"Time: Your next lesson at {actualtime} is in {cellData}");
                            }
                        }
                    }
                }
            }

            public void Start()
            {
                timer.Start();
            } //no need to stop the timer since there will always be a datatable for it to reference (even if its empty)
        }

        public class ToastNotification
        {
            private readonly Popup popup;
            private readonly DispatcherTimer timer;
            public System.Windows.Media.Brush backgroundcolour = new SolidColorBrush(Colors.Blue);
            public System.Windows.Media.Brush textcolour = new SolidColorBrush(Colors.White);


            public ToastNotification()
            {
                popup = new Popup
                {
                    AllowsTransparency = true,
                    Placement = PlacementMode.Absolute,
                    StaysOpen = true
                };

                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(3) //notification lifespan
                };
                timer.Tick += (s, e) =>
                {
                    timer.Stop(); //once the lifespan has finished
                    popup.IsOpen = false;
                };
            }

            public void ShowNotification(string message)
            {
                Border border = new Border
                {
                    Background = backgroundcolour,
                    BorderBrush = new SolidColorBrush(Colors.Black),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(10)
                };

                TextBlock textBlock = new TextBlock
                {
                    Text = message,
                    Foreground = textcolour,
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 600,
                    FontSize = 20,
                    FontWeight = FontWeights.SemiBold,
                };

                border.Child = textBlock;
                popup.Child = border;

                System.Drawing.Rectangle screen = Screen.PrimaryScreen.WorkingArea;

                popup.HorizontalOffset = screen.Right - 320; //screen position (absolute distance from corner rather than fraction)
                popup.VerticalOffset = screen.Bottom - 10;

                popup.IsOpen = true;
                timer.Start();
            }
        }







        //form object code behinds
        //-------------------------
        //------------------------

        private void Run_Astar_Click(object sender, RoutedEventArgs e)
        {
            if (databaseloaded)
            {
                ListBox1.Items.Clear();
                if (string.IsNullOrEmpty(start.Text) || string.IsNullOrEmpty(end.Text))
                {
                    toast.ShowNotification("One or more required boxes ar empty");
                }
                else if (string.IsNullOrEmpty(stop1.Text))
                {
                    foreach (string item in Astar(start.Text, end.Text, blocked.Text)) // displays route
                    {
                        ListBox1.Items.Add(item);
                    }
                }
                else
                {
                    foreach (string item in Astar(start.Text, stop1.Text, blocked.Text))
                    {
                        ListBox1.Items.Add(item);
                    }
                    ListBox1.Items.Add("Reached your stop");
                    foreach (string item in Astar(stop1.Text, end.Text, blocked.Text))
                    {
                        ListBox1.Items.Add(item);
                    }
                }
            }
            else
            {
                toast.ShowNotification("Database not loaded");
            }
        }



        private void GetWeather_Click_1(object sender, RoutedEventArgs e)
        {
            clearweather();
            if (string.IsNullOrEmpty(Location.Text))
            {
                toast.ShowNotification("No input detected, used default location");

                APIRequest(("51.633985", "0.422817"));
            }
            else
            {

                LocationFinder(Location.Text);
            }
        }

        

        private void filepath_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog DbPath = new System.Windows.Forms.OpenFileDialog();
                DbPath.DefaultExt = ".accdb";
                DbPath.Filter = "(.accdb)|*.accdb";
                DbPath.ShowDialog();
                string path = DbPath.FileName;
                path = path.Replace("\"", "");
                connectionstring = connectionstring + path;
                graph = CreateDbGraph();
                fillboxes();
                time = new TimeMonitor(loadtimetable());
                time.Start();
                Timetable.IsReadOnly = false;
                databaseloaded = true;
                info.Content = ""; //the label shown on the timetable screen saying "please load timetable..."
                toast.ShowNotification("Data Loaded successfully!");
            }
            catch (Exception ex)
            {
                toast.ShowNotification("An error was encountered/ File selection cancelled");
            }
        }



        private void committimetable_Click(object sender, RoutedEventArgs e)
        {
            if (databaseloaded)
            {
                if (giveoptions("Are you sure you would like to continue? You are about to override the current information in the database"))
                {
                    DataTable dbupload = new DataTable();
                    foreach (DataGridColumn column in Timetable.Columns)
                    {
                        dbupload.Columns.Add(column.Header.ToString(), typeof(string)); //adds the column headers to the datatable from the datagrid
                    }
                    foreach (Object item in Timetable.Items)
                    {
                        DataRow row = dbupload.NewRow();

                        for (int i = 0; i < Timetable.Columns.Count; i++)
                        {
                            DataGridColumn column = Timetable.Columns[i];
                            string columnName = column.Header.ToString();

                            FrameworkElement cellContent = Timetable.Columns[i].GetCellContent(item);
                            if (cellContent is TextBlock textBlock)
                            {
                                row[columnName] = textBlock.Text;
                            }
                        }

                        dbupload.Rows.Add(row);
                    }
                    using (OleDbConnection connection = new OleDbConnection(connectionstring))
                    {
                        connection.Open();
                        using (OleDbCommand delete = new OleDbCommand($"DELETE FROM Timetable", connection))
                        {
                            delete.ExecuteNonQuery();  // Delete all existing records
                        }

                        foreach (DataRow row in dbupload.Rows)
                        {
                            try
                            {
                                string columnNames = string.Join(", ", dbupload.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"));
                                string paramNames = string.Join(", ", dbupload.Columns.Cast<DataColumn>().Select(c => "?"));
                                string insertQuery = $"INSERT INTO Timetable ({columnNames}) VALUES ({paramNames})"; //query to rewrite in new timetable

                                using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                                {
                                    foreach (DataColumn column in dbupload.Columns)
                                    {
                                        command.Parameters.AddWithValue($"@{column.ColumnName}", row[column] ?? DBNull.Value);
                                    }
                                    command.ExecuteNonQuery();
                                }
                            }
                            catch (Exception ex)
                            {
                                toast.ShowNotification("Error: " + ex.Message);
                            }
                        }
                    }
                }
                else
                {
                    toast.ShowNotification("Upload Cancelled");
                }
                toast.ShowNotification("Successfully uploaded");

            }
            else
            {
                toast.ShowNotification("Database Not Loaded");
            }
        }



        private void resetfilter_Click(object sender, RoutedEventArgs e)
        {
            Teachers.DefaultView.RowFilter = null;
        }
        private void Applyfilter_Click(object sender, RoutedEventArgs e)
        {
            string filter1 = roomdropdown.Text;
            string filter2 = teacherdropdown.Text;
            try
            {
                if (string.IsNullOrWhiteSpace(filter2))
                {
                    Teachers.DefaultView.RowFilter = $"RoomID LIKE '%{filter1}%'";
                }
                else if (string.IsNullOrWhiteSpace(filter1))
                {
                    Teachers.DefaultView.RowFilter = $"Initials LIKE '%{filter2}%'";
                }
                else
                {
                    Teachers.DefaultView.RowFilter = $"RoomID LIKE '%{filter1}%' OR Initials LIKE '%{filter2}%'";
                }
            }
            catch (Exception ex)
            {
                toast.ShowNotification(ex.ToString());
            }

        }



        private void Apply_Click(object sender, RoutedEventArgs e) //changes the colours depending on which boxes are filled
        {
            if (!string.IsNullOrEmpty(backbox.Text))
            {
                System.Windows.Media.Color backcolour = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(backbox.Text);
                this.Background = new SolidColorBrush(backcolour);
                foreach (TabItem tab in Maintabcontrol.Items) //changes colour of every tab within 
                {
                    tab.Background = this.Background;
                }
                foreach (TabItem tab in buildings.Items)
                {
                    tab.Background = this.Background;
                }
                foreach (TabItem tab in building1.Items)
                {
                    tab.Background = this.Background;
                }
                foreach (TabItem tab in building2.Items)
                {
                    tab.Background = this.Background;
                }
                foreach (TabItem tab in building3.Items)
                {
                    tab.Background = this.Background;
                }
                buildings.Background = this.Background;
                building1.Background = this.Background;
                building2.Background = this.Background;
                building3.Background = this.Background;
                grid1.Background = this.Background;
                grid2.Background = this.Background;
                grid3.Background = this.Background;
                grid4.Background = this.Background;
                grid5.Background = this.Background;
                grid6.Background = this.Background;
                grid7.Background = this.Background;
                grid8.Background = this.Background;
                grid9.Background = this.Background;
                grid10.Background = this.Background;
                grid11.Background = this.Background;
                grid13.Background = this.Background;
                grid14.Background = this.Background;
            }
            if (!string.IsNullOrEmpty(notifbackcolour.Text))
            {
                System.Windows.Media.Color backcolour = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(notifbackcolour.Text);
                toast.backgroundcolour = new SolidColorBrush(backcolour);
            }
            if (!string.IsNullOrEmpty(Notiftextcolour.Text))
            {
                System.Windows.Media.Color textcolour = (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(Notiftextcolour.Text);
                toast.textcolour = new SolidColorBrush(textcolour);
            }
            toast.ShowNotification("Test");

        }

        private void download_Click(object sender, RoutedEventArgs e)//allow the user to reload their timetable if they mess up
        {
            if (giveoptions("Are you sure? You are about to overwrite the currently displayed timetable"))
            {
                loadtimetable();
            }
        }
        private void ListBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void TeacherDisplay_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void dropdown1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void dropdown2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void Timetable_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void blocked_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void start_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void end_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void backbox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void textbox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void Location_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        private void Maintabcontrol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Notiftextcolour_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void stop1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void Start_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void End_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ListBox_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {

        }

        private void building2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}