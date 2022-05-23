using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using PartsInventoryConnector.Authentication;
using PartsInventoryConnector.Console;
using PartsInventoryConnector.Data;
using PartsInventoryConnector.MicrosoftGraph;
using PartsInventoryConnector.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PartsInventoryConnector
{
    class Program
    {
        private static MicrosoftGraphHelper _microsoftGraphHelper;

        private static ExternalConnection _currentConnection;

        private static string _tenantId;

        static async Task Main(string[] args)
        {
            try
            {
                // Load configuration from appsettings.json
                var appConfig = LoadAppSettings();
                if (appConfig == null)
                {
                    return;
                }

                // Save tenant ID for setting ACL on items
                _tenantId = appConfig["tenantId"];

                // Initialize the auth provider
                var authProvider = new ClientCredentialAuthProvider(
                    appConfig["appId"],
                    appConfig["tenantId"],
                    appConfig["appSecret"]
                );

                // Check if the database is empty
                using (var db = new ApplianceDbContext())
                {
                    if (db.Parts.IgnoreQueryFilters().Count() <= 0)
                    {
                        ImportCsvToDatabase(db, "ApplianceParts.csv");
                    }
                }

                _microsoftGraphHelper = new MicrosoftGraphHelper(authProvider);

                do
                {
                    var userChoice = DoMenuPrompt();

                    switch (userChoice)
                    {
                        case MenuChoice.CreateConnection:
                            await CreateConnectionAsync();
                            break;
                        case MenuChoice.RegisterSchema:
                            await RegisterSchemaAsync();
                            break;
                        case MenuChoice.PushAllItems:
                            await UpdateItemsFromDatabase();
                            break;
                        case MenuChoice.Exit:
                            // Exit the program
                            return;
                        case MenuChoice.Invalid:
                        default:
                            break;
                    }

                } while (true);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }

        private static async Task UpdateItemsFromDatabase()
        {
            if (_currentConnection == null)
            {
                System.Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            List<AppliancePart> partsToUpload = null;

            using (var db = new ApplianceDbContext())
            {
                partsToUpload = db.Parts.ToList();
            }

            System.Console.WriteLine($"Processing {partsToUpload.Count()} add/updates");

            foreach (var part in partsToUpload)
            {
                var newItem = new ExternalItem
                {
                    Id = part.PartNumber.ToString(),
                    Content = new ExternalItemContent
                    {
                        // Need to set to null, service returns 400
                        // if @odata.type property is sent
                        ODataType = null,
                        Type = ExternalItemContentType.Text,
                        Value = part.Description
                    },
                    Acl = new List<Acl>
                    {
                        new Acl
                        {
                            AccessType = AccessType.Grant,
                            Type = AclType.Everyone,
                            Value = _tenantId,
                            IdentitySource = "Azure Active Directory"
                        }
                    },
                    Properties = part.AsExternalItemProperties()
                };

                try
                {
                    System.Console.Write($"Uploading part number {part.PartNumber}...");
                    await _microsoftGraphHelper.AddOrUpdateItem(_currentConnection.Id, newItem);
                    System.Console.WriteLine("DONE");
                }
                catch (ServiceException serviceException)
                {
                    System.Console.WriteLine("FAILED");
                    System.Console.WriteLine($"{serviceException.StatusCode} error adding or updating part {part.PartNumber}");
                    System.Console.WriteLine(serviceException.Message);
                }
            }
        }

        private static async Task RegisterSchemaAsync()
        {
            if (_currentConnection == null)
            {
                System.Console.WriteLine("No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            System.Console.WriteLine("Registering schema, this may take a moment...");

            try
            {
                // Register the schema
                var schema = new Schema
                {
                    // Need to set to null, service returns 400
                    // if @odata.type property is sent
                    ODataType = null,
                    BaseType = "microsoft.graph.externalItem",
                    Properties = new List<Property>
                    {
                        new Property { Name = "partNumber", Type = PropertyType.Int64, IsQueryable = true, IsSearchable = false, IsRetrievable = true },
                        new Property { Name = "name", Type = PropertyType.String, IsQueryable = true, IsSearchable = true, IsRetrievable = true },
                        new Property { Name = "description", Type = PropertyType.String, IsQueryable = false, IsSearchable = true, IsRetrievable = true },
                        new Property { Name = "price", Type = PropertyType.Double, IsQueryable = true, IsSearchable = false, IsRetrievable = true },
                        new Property { Name = "inventory", Type = PropertyType.Int64, IsQueryable = true, IsSearchable = false, IsRetrievable = true },
                        new Property { Name = "appliances", Type = PropertyType.StringCollection, IsQueryable = true, IsSearchable = true, IsRetrievable = true }
                    }
                };

                await _microsoftGraphHelper.RegisterSchemaAsync(_currentConnection.Id, schema);
                System.Console.WriteLine("Schema registered");
            }
            catch (ServiceException serviceException)
            {
                System.Console.WriteLine($"{serviceException.StatusCode} error registering schema:");
                System.Console.WriteLine(serviceException.Message);
                return;
            }
        }

        private static async Task CreateConnectionAsync()
        {
            var connectionId = PromptForInput("Enter a unique ID for the new connection", true);
            var connectionName = PromptForInput("Enter a name for the new connection", true);
            var connectionDescription = PromptForInput("Enter a description for the new connection", false);

            try
            {
                // Create the connection
                _currentConnection = await _microsoftGraphHelper.CreateConnectionAsync(connectionId, connectionName, connectionDescription);
                System.Console.WriteLine("New connection created");
                System.Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(_currentConnection, Newtonsoft.Json.Formatting.Indented));
            }
            catch (ServiceException serviceException)
            {
                System.Console.WriteLine(serviceException.Message);
                return;
            }
        }

        private static void ImportCsvToDatabase(ApplianceDbContext db, string partsFilePath)
        {
            var parts = CsvDataLoader.LoadPartsFromCsv(partsFilePath);
            db.AddRange(parts);
            db.SaveChanges();
        }

        private static MenuChoice DoMenuPrompt()
        {
            System.Console.WriteLine($"Current connection: {(_currentConnection == null ? "NONE" : _currentConnection.Name)}");
            System.Console.WriteLine("Please choose one of the following options:");

            System.Console.WriteLine($"{Convert.ToInt32(MenuChoice.CreateConnection)}. Create a connection");
            System.Console.WriteLine($"{Convert.ToInt32(MenuChoice.RegisterSchema)}. Register schema for current connection");
            System.Console.WriteLine($"{Convert.ToInt32(MenuChoice.PushAllItems)}. Push ALL items to current connection");
            System.Console.WriteLine($"{Convert.ToInt32(MenuChoice.Exit)}. Exit");

            try
            {
                var choice = int.Parse(System.Console.ReadLine());
                return (MenuChoice)choice;
            }
            catch (FormatException)
            {
                return MenuChoice.Invalid;
            }
        }
        private static string PromptForInput(string prompt, bool valueRequired)
        {
            string response = null;

            do
            {
                System.Console.WriteLine($"{prompt}:");
                response = System.Console.ReadLine();
                if (valueRequired && string.IsNullOrEmpty(response))
                {
                    System.Console.WriteLine("You must provide a value");
                }
            } while (valueRequired && string.IsNullOrEmpty(response));

            return response;
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["appSecret"]) ||
                string.IsNullOrEmpty(appConfig["tenantId"]))
            {
                return null;
            }

            return appConfig;
        }
    }
}